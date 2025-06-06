# app.py

import streamlit as st
import pandas as pd
import re
import PyPDF2
import io
import subprocess
import tempfile
import os

st.set_page_config(page_title="PDF → Excel", layout="wide")
st.title("PDF → Excel")

st.markdown(
    """
    Wgraj plik PDF ze zamówieniem. Aplikacja działa tak:
    1. Próbujemy odczytać tekst przez PyPDF2.
    2. Jeżeli nie znajdzie ani jednej sensownej linii, uruchamiamy OCR za pomocą zainstalowanego w systemie programu `tesseract`.
       - W Streamlit Cloud najczęściej jest już dostępny `tesseract`.  
       - Konwertujemy każdą stronę PDF na pojedynczy obraz PNG (narzędziem `pdftoppm`, które jest częścią paczki `poppler-utils`),
         a potem wywołujemy `tesseract --oem 1 --psm 3 <image>.png stdout -l pol`, żeby otrzymać tekst.
    3. Gdy mamy już listę wierszy tekstu (`all_lines`), wykrywamy układ:
       - **Układ B**: wiersz pojedynczy, w formie  
         ```
         <Lp> <EAN(13)> <pełna nazwa> <ilość>,<xx> szt. ...
         ```  
         – wyciągamy te cztery pola (Lp, Barcode, Name, Quantity).
       - **Układ A**: Lp to samodzielna linia z liczbą, „Kod kres.: <EAN>” jest w osobnej linii, a nazwa może być rozbita przed i po kolumnie cen.  
       – Stary parser dla Układu A przypisuje każdemu Lp kod EAN z najbliższej linii „Kod kres” i scala fragmenty nazwy przed i po kolumnie cen.
    4. Wyświetlamy wynik jako tabelę i umożliwiamy pobranie pliku Excel.
    """
)

# ──────────────────────────────────────────────────────────────────────────────

def extract_text_with_py(pdf_bytes: bytes) -> list[str]:
    """
    Wyciąga każdą niepustą linię tekstu przez PyPDF2.
    Zwraca listę wierszy; jeśli jest pusta, oznacza, że PyPDF2 nie znalazł tekstu.
    """
    reader = PyPDF2.PdfReader(io.BytesIO(pdf_bytes))
    lines = []
    for page in reader.pages:
        text = page.extract_text() or ""
        for ln in text.split("\n"):
            stripped = ln.strip()
            if stripped:
                lines.append(stripped)
    return lines


def ocr_pdf_to_lines(pdf_bytes: bytes) -> list[str]:
    """
    Jeżeli PDF nie zawiera osadzonego tekstu, konwertujemy każdą stronę na obraz PNG
    (poprzez narzędzie `pdftoppm`) i robimy OCR przez `tesseract`.
    Zwracamy listę wierszy.
    """
    lines = []
    with tempfile.TemporaryDirectory() as tmpdir:
        # Zapiszmy PDF do pliku tymczasowego
        tmp_pdf = os.path.join(tmpdir, "temp.pdf")
        with open(tmp_pdf, "wb") as f:
            f.write(pdf_bytes)
        # Użyjemy pdftoppm, by prześwietlić strony jako PNG-y
        # Nazwa plików wyjściowych: page-1.png, page-2.png, ...
        cmd = ["pdftoppm", "-png", "-r", "300", tmp_pdf, os.path.join(tmpdir, "page")]
        try:
            subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception:
            return []  # pdftoppm może nie być w systemie
        # Teraz każdy plik page-1.png, page-2.png, ...
        idx = 1
        while True:
            img_path = os.path.join(tmpdir, f"page-{idx}.png")
            if not os.path.exists(img_path):
                break
            # OCR przez tesseract, język polski (język można zmienić, jeśli nie ma 'pol')
            try:
                result = subprocess.run(
                    ["tesseract", img_path, "stdout", "-l", "pol"],
                    check=True, capture_output=True
                )
                text = result.stdout.decode("utf-8", errors="ignore")
            except Exception:
                text = ""
            for ln in text.split("\n"):
                stripped = ln.strip()
                if stripped:
                    lines.append(stripped)
            idx += 1
    return lines


def extract_text_from_pdf_bytes(pdf_bytes: bytes) -> list[str]:
    """
    Najpierw próbujemy PyPDF2. Jeśli nic nie znajdzie, przechodzimy do OCR.
    """
    lines = extract_text_with_py(pdf_bytes)
    if lines:
        return lines
    # inaczej OCR:
    ocr_lines = ocr_pdf_to_lines(pdf_bytes)
    return ocr_lines


def parse_layout_b(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla układu z EAN w tej samej linii, np.
    1 5029040012366 Nazwa Produktu ... 96,00 szt. ...
    """
    products = []
    pattern = re.compile(
        r"^(\d+)\s+(\d{13})\s+(.+?)\s+(\d{1,3}),\d{2}\s+szt",
        flags=re.IGNORECASE
    )
    for ln in all_lines:
        m = pattern.match(ln)
        if m:
            Lp_val = int(m.group(1))
            Barcode_val = m.group(2)
            Name_val = m.group(3).strip()
            Quantity_val = int(m.group(4).replace(" ", ""))
            products.append({
                "Lp": Lp_val,
                "Name": Name_val,
                "Quantity": Quantity_val,
                "Barcode": Barcode_val
            })
    return pd.DataFrame(products)


def parse_layout_a(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla klasycznego układu „Kod kres.: <EAN> w oddzielnej linii”.
    """
    idx_lp = []
    for i in range(len(all_lines) - 1):
        if re.fullmatch(r"\d+", all_lines[i]):
            nxt = all_lines[i + 1]
            if (
                re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", nxt)
                and nxt.lower() != "szt."
                and not re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", nxt)
                and not nxt.startswith("Kod kres")
            ):
                idx_lp.append(i)

    idx_ean = [i for i, ln in enumerate(all_lines) if ln.startswith("Kod kres")]

    products = []
    for idx, lp_idx in enumerate(idx_lp):
        prev_lp = idx_lp[idx - 1] if idx > 0 else -1
        next_lp = idx_lp[idx + 1] if idx + 1 < len(idx_lp) else len(all_lines)

        # 1) EAN
        barcode = None
        valid_eans = [e for e in idx_ean if prev_lp < e < next_lp]
        if valid_eans:
            eidx = max(valid_eans)
            parts = all_lines[eidx].split(":", 1)
            if len(parts) == 2:
                barcode = parts[1].strip()

        # 2) Nazwa + Quantity
        name_parts = []
        qty = None
        qty_idx = None

        for j in range(lp_idx + 1, next_lp):
            ln = all_lines[j]
            if re.fullmatch(r"\d+", ln) and (j + 1 < next_lp and all_lines[j + 1].lower() == "szt."):
                qty_idx = j
                qty = int(ln)
                break
            if (
                re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", ln)
                and not re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", ln)
                and not ln.startswith("VAT")
                and ln != "/"
                and not ln.startswith("ARA")
                and not ln.startswith("KAT")
            ):
                name_parts.append(ln)

        if qty_idx is None:
            continue

        for k in range(qty_idx + 1, next_lp):
            ln2 = all_lines[k]
            if ln2.startswith("Kod kres"):
                break
            if (
                re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", ln2)
                and not re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", ln2)
                and not ln2.startswith("VAT")
                and ln2 != "/"
                and not ln2.startswith("ARA")
                and not ln2.startswith("KAT")
            ):
                name_parts.append(ln2)

        Name_val = " ".join(name_parts).strip()
        products.append({
            "Lp": int(all_lines[lp_idx]),
            "Name": Name_val,
            "Quantity": qty,
            "Barcode": barcode
        })

    return pd.DataFrame(products)


# ──────────────────────────────────────────────────────────────────────────────

# 1) FileUploader – wgraj PDF
uploaded_file = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if uploaded_file is None:
    st.info("Proszę wgrać plik PDF, aby kontynuować.")
    st.stop()

# 2) Pobierz bajty PDF-a
pdf_bytes = uploaded_file.read()

# 3) Wyciągnij wszystkie linie tekstu (PyPDF2 albo OCR)
all_lines = extract_text_from_pdf_bytes(pdf_bytes)

# 4) Jeżeli nadal brak linii → wyświetl komunikat i zakończ
if not all_lines:
    st.error(
        "Nie udało się wyciągnąć tekstu z PDF-a. "
        "Prawdopodobnie jest to czysty skan/obraz. "
        "Aby go sparsować, najpierw wykonaj OCR (np. `tesseract`) lokalnie."
    )
    st.stop()

# 5) Wykryj układ B: Lp + EAN w jednej linii
pattern_b = re.compile(
    r"^\d+\s+\d{13}\s+.+\s+\d{1,3},\d{2}\s+szt", flags=re.IGNORECASE
)
is_layout_b = any(pattern_b.match(ln) for ln in all_lines)

# 6) Parsowanie w zależności od układu
if is_layout_b:
    df = parse_layout_b(all_lines)
else:
    df = parse_layout_a(all_lines)

# 7) Jeśli kolumny "Name" i "Quantity" istnieją, odfiltruj brakujące
if "Name" in df.columns and "Quantity" in df.columns:
    df = df.dropna(subset=["Name", "Quantity"]).reset_index(drop=True)

# 8) Wyświetl wynik
st.subheader("Wyekstrahowane pozycje zamówienia")
st.dataframe(df, use_container_width=True)

# 9) Przycisk do pobrania Excel
def convert_df_to_excel(df_in: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_in.to_excel(writer, index=False, sheet_name="Zamówienie")
    return output.getvalue()

excel_data = convert_df_to_excel(df)
st.download_button(
    label="Pobierz wynik jako Excel",
    data=excel_data,
    file_name="parsed_zamowienie.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
