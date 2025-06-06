# app.py

import streamlit as st
import pandas as pd
import re
import fitz  # PyMuPDF
import PyPDF2
import io
import subprocess
import tempfile
import os

st.set_page_config(page_title="PDF → Excel", layout="wide")
st.title("PDF → Excel")

st.markdown(
    """
    Wgraj plik PDF ze zamówieniem. Aplikacja:
    1. Próbuje wyciągnąć tekst przez PyMuPDF (fitz).  
       Jeśli nie znajdzie ani jednej czytelnej linii, próbuje PyPDF2.  
       Jeśli dalej nic nie ma – przeprowadza OCR przez `pdftoppm` + `tesseract`.  
    2. Mając listę wierszy tekstu (`all_lines`), wykrywa układ:
       - **Układ B**: każda pozycja w jednej linii (np. 
         `1 5029040012366 Nazwa Produktu 96,00 szt.`).  
       - **Układ C**: EAN w osobnej linii (tylko 13 cyfr), potem `Lp`, potem `Name`, itd.  
       - **Układ A**: „Kod kres.: <EAN>” w osobnej linii, Lp w osobnej linii (czysta liczba), 
         fragmenty nazwy przed i po kolumnie cen.  
    3. W zależności od wykrytego układu wywołuje odpowiedni parser (A, B lub C).  
    4. Wyświetla tabelę z kolumnami: `Lp`, `Name`, `Quantity`, `Barcode`.  
    5. Umożliwia pobranie wynikowego pliku Excel.
    """
)

# ──────────────────────────────────────────────────────────────────────────────

def extract_text_with_fitz(pdf_bytes: bytes) -> list[str]:
    """
    Próba wyciągnięcia tekstu każdej strony przez PyMuPDF (fitz).
    Zwraca listę niepustych wierszy. Jeśli niczego nie znajdzie, zwraca [].
    """
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception:
        return []
    lines = []
    for page in doc:
        text = page.get_text() or ""
        for ln in text.split("\n"):
            stripped = ln.strip()
            if stripped:
                lines.append(stripped)
    return lines


def extract_text_with_pypdf2(pdf_bytes: bytes) -> list[str]:
    """
    Próba wyciągnięcia tekstu każdej strony przez PyPDF2.
    Zwraca listę niepustych wierszy. Jeśli niczego nie znajdzie, zwraca [].
    """
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(pdf_bytes))
    except Exception:
        return []
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
    Jeśli PDF nie zawiera osadzonego tekstu, wykonaj OCR:
    1) Zapisz bajty do pliku tymczasowego.
    2) Użyj `pdftoppm` (poppler-utils) do wygenerowania PNG stron.
    3) Na każdym PNG uruchom `tesseract ... stdout -l pol`, by pozyskać tekst.
    4) Zwróć listę niepustych wierszy.
    """
    lines = []
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_pdf = os.path.join(tmpdir, "temp.pdf")
        with open(tmp_pdf, "wb") as f:
            f.write(pdf_bytes)
        # Konwertuj każdą stronę PDF na PNG: page-1.png, page-2.png, ...
        cmd = ["pdftoppm", "-png", "-r", "300", tmp_pdf, os.path.join(tmpdir, "page")]
        try:
            subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception:
            return []
        idx = 1
        while True:
            img_path = os.path.join(tmpdir, f"page-{idx}.png")
            if not os.path.exists(img_path):
                break
            try:
                result = subprocess.run(
                    ["tesseract", img_path, "stdout", "-l", "pol"],
                    capture_output=True, check=True
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
    Łączy wszystkie podejścia:
    1) Najpierw PyMuPDF (fitz),
    2) potem PyPDF2,
    3) i w końcu OCR (pdftoppm + tesseract).
    Jeśli wciąż brak linii, zwraca [].
    """
    lines = extract_text_with_fitz(pdf_bytes)
    if lines:
        return lines
    lines = extract_text_with_pypdf2(pdf_bytes)
    if lines:
        return lines
    return ocr_pdf_to_lines(pdf_bytes)


def parse_layout_b(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla układu B: każda pozycja w jednej linii, np.
    1 5029040012366 Nazwa Produktu 96,00 szt. ...
    -> wyciągamy Lp (grupa 1), Barcode (grupa 2), Name (grupa 3), Quantity (grupa 4).
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


def parse_layout_c(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla układu C: EAN w osobnej linii (13 cyfr),
    potem Lp, potem Name, potem "ilość" w następującej strukturze:
      <EAN (13-digit)>
      <Lp>               <-- pure integer
      <Name>
      <Price>
      "szt."
      <VAT or 0,00>
      <Quantity>         <-- pure integer
      ...
    Mapujemy każdy Lp do EAN najbliższego powyżej (e_idx < lp_idx).
    Name to linia bezpośrednio po Lp. Quantity to the integer two lines after "szt.".
    """
    # 1) Znajdź indeksy Lp (linia czysta-liczba, pod nią coś z literami)
    idx_lp = []
    for i in range(len(all_lines) - 1):
        if re.fullmatch(r"\d+", all_lines[i]):
            nxt = all_lines[i + 1]
            if re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", nxt):
                idx_lp.append(i)

    # 2) Znajdź indeksy czystych 13-cyfrowych EAN-ów
    idx_ean = [i for i, ln in enumerate(all_lines) if re.fullmatch(r"\d{13}", ln)]

    products = []
    for idx, lp_idx in enumerate(idx_lp):
        # 2a) EAN: spośród e_idx < lp_idx wybierz największy
        eans = [e for e in idx_ean if e < lp_idx]
        barcode = None
        if eans:
            barcode = all_lines[max(eans)]

        # 2b) Name to linia immediately after lp_idx
        Name_val = all_lines[lp_idx + 1]

        # 2c) Quantity: find "szt." below lp_idx, then qty = integer at +2
        qty = None
        for j in range(lp_idx + 1, len(all_lines) - 2):
            if all_lines[j].lower() == "szt.":
                if re.fullmatch(r"\d+", all_lines[j + 2]):
                    qty = int(all_lines[j + 2])
                break

        # Jeżeli nie znaleziono qty, pomijamy tę pozycję
        if qty is None:
            continue

        products.append({
            "Lp": int(all_lines[lp_idx]),
            "Name": Name_val.strip(),
            "Quantity": qty,
            "Barcode": barcode
        })

    return pd.DataFrame(products)


def parse_layout_a(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla układu A: "Kod kres.: <EAN>" w oddzielnej linii,
    Lp to linia czysta-liczba, pod nią początek fragmentu nazwy,
    nazwa może się dzielić przed/po kolumnie cen/ilości.
    """
    # 1) Znajdź indeksy Lp (linia czysta-liczba, pod nią linia z literami)
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

    # 2) Znajdź indeksy linii "Kod kres"
    idx_ean = [i for i, ln in enumerate(all_lines) if ln.startswith("Kod kres")]

    products = []
    for idx, lp_idx in enumerate(idx_lp):
        prev_lp = idx_lp[idx - 1] if idx > 0 else -1
        next_lp = idx_lp[idx + 1] if idx + 1 < len(idx_lp) else len(all_lines)

        # 2a) EAN: spośród e in idx_ean takich, że prev_lp < e < next_lp, wybierz maksymalny
        e_subset = [e for e in idx_ean if prev_lp < e < next_lp]
        barcode = None
        if e_subset:
            parts = all_lines[max(e_subset)].split(":", 1)
            if len(parts) == 2:
                barcode = parts[1].strip()

        # 2b) Nazwa + Ilość:
        name_parts = []
        qty = None
        qty_idx = None

        # Fragmenty nazwy przed kolumną ilości/ceny
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

        # Jeżeli brak qty_idx → pomiń
        if qty_idx is None:
            continue

        # Po znalezieniu ilości, zbieraj kolejne fragmenty nazwy aż do "Kod kres"
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

# 1) FileUploader – użytkownik wgrywa PDF
uploaded_file = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if uploaded_file is None:
    st.info("Proszę wgrać plik PDF, aby kontynuować.")
    st.stop()

# 2) Pobierz bajty PDF-a
pdf_bytes = uploaded_file.read()

# 3) Wyciągnij wszystkie linie tekstu (fitz → PyPDF2 → OCR)
all_lines = extract_text_from_pdf_bytes(pdf_bytes)

# 4) Jeśli nadal brak linii – zakończ z komunikatem
if not all_lines:
    st.error(
        "Nie udało się wyciągnąć tekstu z PDF-a. "
        "Prawdopodobnie jest to czysty skan/obraz i nie udał się OCR. "
        "Upewnij się, że `pdftoppm` i `tesseract` są zainstalowane."
    )
    st.stop()

# 5) Wykryj układ B (Lp + EAN w tej samej linii)
pattern_b = re.compile(r"^\d+\s+\d{13}\s+.+\s+\d{1,3},\d{2}\s+szt", flags=re.IGNORECASE)
is_layout_b = any(pattern_b.match(ln) for ln in all_lines)

# 6) Wykryj układ C: jeżeli jest czysty 13-cyfrowy EAN w osobnej linii,
#    ale nie pasuje do układu B
has_pure_ean = any(re.fullmatch(r"\d{13}", ln) for ln in all_lines)
is_layout_c = has_pure_ean and not is_layout_b

# 7) Parsuj w zależności od układu
if is_layout_b:
    df = parse_layout_b(all_lines)
elif is_layout_c:
    df = parse_layout_c(all_lines)
else:
    df = parse_layout_a(all_lines)

# 8) Jeżeli w DataFrame są kolumny "Name" i "Quantity", usuń wiersze z brakami
if "Name" in df.columns and "Quantity" in df.columns:
    df = df.dropna(subset=["Name", "Quantity"]).reset_index(drop=True)

# 9) Wyświetl wynik w Streamlit
st.subheader("Wyekstrahowane pozycje zamówienia")
st.dataframe(df, use_container_width=True)

# 10) Przygotuj przycisk do pobrania pliku Excel
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
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
