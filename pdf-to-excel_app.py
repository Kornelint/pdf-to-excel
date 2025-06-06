# app.py

import streamlit as st
import pandas as pd
import re
import io
import subprocess
import tempfile
import os

from pdfminer.high_level import extract_text as pdfminer_extract_text
import PyPDF2

st.set_page_config(page_title="PDF → Excel", layout="wide")
st.title("PDF → Excel")

st.markdown(
    """
    Wgraj plik PDF ze zamówieniem. Aplikacja:
    1. Próbuje wyciągnąć tekst w trzech krokach:
       a) **pdfminer.six** – najlepszy do wielu „dziwnych” PDF-ów.
       b) **PyPDF2** – jeśli pdfminer zwróci pusty wynik.
       c) **OCR** (`pdftoppm` + `tesseract`) – jeśli dalej brak czytelnych linii.
    2. Gdy mamy listę wierszy tekstu (`all_lines`), wykrywa układ:
       - **Układ B**: każda pozycja w jednej linii, np.  
         `1 5029040012366 Nazwa Produktu 96,00 szt.`  
       - **Układ C**: czysty 13-cyfrowy EAN w osobnej linii, potem Lp, potem Name, potem `szt.` i Quantity.  
       - **Układ A**: „Kod kres.: <EAN>” w oddzielnej linii, Lp w oddzielnej linii (czysta liczba), 
         fragmenty nazwy przed i po kolumnie cen.
    3. Parsuje odpowiednio (A, B lub C), wyświetla tabelę z kolumnami:
       `Lp`, `Name`, `Quantity`, `Barcode` oraz pozwala pobrać plik Excel.
    """
)

# ──────────────────────────────────────────────────────────────────────────────

def extract_with_pdfminer(pdf_bytes: bytes) -> list[str]:
    """
    Próbuj wydobyć tekst każdej strony z pdfminer.six. 
    Zwraca listę wszystkich niepustych wierszy. 
    Jeśli nic się nie wydobędzie, zwraca pustą listę.
    """
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(pdf_bytes)
        tmp_path = tmp.name

    try:
        text = pdfminer_extract_text(tmp_path)
    except Exception:
        os.unlink(tmp_path)
        return []
    os.unlink(tmp_path)

    lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
    return lines


def extract_with_pypdf2(pdf_bytes: bytes) -> list[str]:
    """
    Próbuj wyciągnąć tekst każdej strony przez PyPDF2.
    Zwraca listę niepustych wierszy. Jeśli nie znajdzie niczego, zwraca [].
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
    Jeśli wcześniej nie udało się wyciągnąć tekstu, wykonuje OCR (przez pdftoppm + tesseract).
    Zwraca listę niepustych wierszy. Jeśli nie uda się, zwraca [].
    """
    lines = []
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_pdf = os.path.join(tmpdir, "temp.pdf")
        with open(tmp_pdf, "wb") as f:
            f.write(pdf_bytes)

        # pdftoppm → PNG-y stron (300 DPI)
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
    Łączy wszystkie metody ekstrakcji:
    1) pdfminer.six
    2) PyPDF2
    3) OCR (pdftoppm + tesseract)
    Zwraca listę niepustych linii. Jeśli wciąż nic nie ma, zwraca [].
    """
    lines = extract_with_pdfminer(pdf_bytes)
    if lines:
        return lines

    lines = extract_with_pypdf2(pdf_bytes)
    if lines:
        return lines

    return ocr_pdf_to_lines(pdf_bytes)


def parse_layout_b(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla układu B – każda pozycja w jednej linii:
      <Lp> <EAN(13)> <pełna nazwa> <ilość>,<xx> szt ...
    Wyciąga Lp, Barcode, Name, Quantity.
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
    Parser dla układu C – czysty 13-cyfrowy EAN w oddzielnej linii, potem Lp, potem Name, potem "szt." i Quantity.
    Logika:
      1) Znajdź wszystkie indeksy Lp: linia tylko liczba, pod nią fragment nazwy.
      2) Znajdź wszystkie indeksy czystych 13-cyfrowych EAN-ów.
      3) Dla każdego Lp przypisz EAN z maksymalnego e < lp_idx.
      4) Name = all_lines[lp_idx + 1].
      5) Quantity = integer dwie linie po napotkaniu "szt." poniżej lp_idx.
    """
    idx_lp = []
    for i in range(len(all_lines) - 1):
        if re.fullmatch(r"\d+", all_lines[i]):
            nxt = all_lines[i + 1]
            if re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", nxt):
                idx_lp.append(i)

    idx_ean = [i for i, ln in enumerate(all_lines) if re.fullmatch(r"\d{13}", ln)]
    products = []
    for idx, lp_idx in enumerate(idx_lp):
        eans = [e for e in idx_ean if e < lp_idx]
        barcode = all_lines[max(eans)] if eans else None

        Name_val = all_lines[lp_idx + 1] if lp_idx + 1 < len(all_lines) else None
        qty = None
        for j in range(lp_idx + 1, len(all_lines) - 2):
            if all_lines[j].lower() == "szt." and re.fullmatch(r"\d+", all_lines[j + 2]):
                qty = int(all_lines[j + 2])
                break

        if Name_val and qty is not None:
            products.append({
                "Lp": int(all_lines[lp_idx]),
                "Name": Name_val.strip(),
                "Quantity": qty,
                "Barcode": barcode
            })
    return pd.DataFrame(products)


def parse_layout_a(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla układu A – „Kod kres.: <EAN>” w oddzielnej linii,
    Lp w oddzielnej linii (czysta liczba), fragmenty nazwy przed i po kolumnie cen.
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

        valid_eans = [e for e in idx_ean if prev_lp < e < next_lp]
        barcode = None
        if valid_eans:
            parts = all_lines[max(valid_eans)].split(":", 1)
            if len(parts) == 2:
                barcode = parts[1].strip()

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

# 3) Ekstrakcja tekstu: pdfminer → PyPDF2 → OCR
all_lines = extract_text_from_pdf_bytes(pdf_bytes)

# 4) Jeśli wciąż brak linii, komunikat i koniec
if not all_lines:
    st.error(
        "Nie udało się wyciągnąć tekstu z PDF (nawet po próbach OCR). "
        "Upewnij się, że PDF ma warstwę tekstową lub wykonaj OCR zewnętrznie."
    )
    st.stop()

# 5) Detekcja układu B: Lp + EAN w jednej linii
pattern_b = re.compile(r"^\d+\s+\d{13}\s+.+\s+\d{1,3},\d{2}\s+szt", flags=re.IGNORECASE)
is_layout_b = any(pattern_b.match(ln) for ln in all_lines)

# 6) Detekcja układu C: czysty 13-cyfrowy EAN w linii, ale nie układ B
has_pure_ean = any(re.fullmatch(r"\d{13}", ln) for ln in all_lines)
is_layout_c = has_pure_ean and not is_layout_b

# 7) Parsowanie w zależności od układu
if is_layout_b:
    df = parse_layout_b(all_lines)
elif is_layout_c:
    df = parse_layout_c(all_lines)
else:
    df = parse_layout_a(all_lines)

# 8) Odfiltrowanie wierszy bez nazwy lub ilości, jeśli kolumny istnieją
if "Name" in df.columns and "Quantity" in df.columns:
    df = df.dropna(subset=["Name", "Quantity"]).reset_index(drop=True)

# 9) Wyświetlenie w Streamlit
st.subheader("Wyekstrahowane pozycje zamówienia")
st.dataframe(df, use_container_width=True)

# 10) Przycisk do pobrania pliku Excel
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
