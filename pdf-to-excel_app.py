# app.py

import streamlit as st
import pandas as pd
import re
import PyPDF2
import io

# Próbujemy zaimportować pdfplumber. Jeśli nie jest zainstalowane, ustawiamy flagę na False.
try:
    import pdfplumber
    _HAS_PDFPLUMBER = True
except ImportError:
    _HAS_PDFPLUMBER = False

st.set_page_config(page_title="PDF → Excel", layout="wide")
st.title("PDF → Excel")

if not _HAS_PDFPLUMBER:
    st.warning(
        "Uwaga: brak biblioteki `pdfplumber`. Bez niej aplikacja może nie poradzić sobie "
        "z niektórymi PDF-ami (np. „Wydruk.pdf”).\n\n"
        "Aby zainstalować, dodaj do swojego `requirements.txt`:\n"
        "`pdfplumber`\n\n"
        "lub w terminalu uruchom:\n"
        "`pip install pdfplumber`\n"
    )

st.markdown(
    """
    Wgraj plik PDF ze zamówieniem. Aplikacja:
    1. Próbuje wyciągnąć tekst przez PyPDF2.
    2. Jeśli PyPDF2 nie zwróci sensownych linii (np. zwróci gibberish), 
       a biblioteka `pdfplumber` jest dostępna (`_HAS_PDFPLUMBER=True`), 
       próbuje wtedy wydobyć tekst z pomocą `pdfplumber`.
    3. Wydobycie zwraca listę niepustych linii. Na ich podstawie:
       - **Układ A**: linia ze słowami „Kod kres.: <EAN>”, potem w kolejnych liniach 
         numer Lp (czysta liczba), fragmenty nazwy przed i po sekcji z „<ilość> szt.”.
       - **Układ B**: każda pozycja w jednej linii, np. `1 5029040012366 Nazwa Produktu 96,00 szt.`  
         Rozbijamy to regexem.
       - **Układ C**: czysty wiersz z 13-cyfrowym EAN, potem oddzielnie Lp, nazwa, „szt.” i ilość.
    4. W zależności od wykrytego układu wywołujemy odpowiedni parser (A, B lub C).
    5. Wyświetlamy tabelę z kolumnami `Lp`, `Name`, `Quantity`, `Barcode` i umożliwiamy pobranie pliku Excel.
    """
)

# ──────────────────────────────────────────────────────────────────────────────

def extract_text_with_pypdf2(pdf_bytes: bytes) -> list[str]:
    """
    Najpierw próbuje wyciągnąć wszystkie linie przez PyPDF2.
    Jeśli PyPDF2 zwróci puste albo „gibberish” (np. pojedyncze znaki), 
    a do dyspozycji mamy pdfplumber, próbuje dalej z niego.
    Jeśli wciąż nic się nie znajdzie, zwraca pustą listę.
    """
    lines = []

    # ---- 1) próba z PyPDF2 ----
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(pdf_bytes))
        for page in reader.pages:
            text = page.extract_text() or ""
            for ln in text.split("\n"):
                stripped = ln.strip()
                if stripped:
                    lines.append(stripped)
    except Exception:
        lines = []

    # Sprawdzamy, czy linie z PyPDF2 wyglądają na sensowne:
    # Szukamy chociaż jednej linii, która jest albo nagłówkiem "Lp" albo czystym 13-cyfrowym EAN-em.
    has_ean_or_header = any(
        re.fullmatch(r"\d{13}", ln) or ln.lower().startswith("lp")
        for ln in lines
    )
    if has_ean_or_header and lines:
        return lines

    # ---- 2) jeśli PyPDF2 nie dał nic sensownego, a mamy pdfplumber, próbujemy z pdfplumber ----
    if _HAS_PDFPLUMBER:
        pl_lines = []
        try:
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                for page in pdf.pages:
                    text = page.extract_text() or ""
                    for ln in text.split("\n"):
                        stripped = ln.strip()
                        if stripped:
                            pl_lines.append(stripped)
        except Exception:
            pl_lines = []

        # Jeśli pdfplumber zwróciło jakieś linie, używamy ich
        if pl_lines:
            return pl_lines

    # Jeśli nie udało się nic wydobyć (albo nie mamy pdfplumber), zwracamy to, co mamy (a może być puste).
    return lines


def parse_layout_a(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla układu A:
    - Gdzieś w all_lines jest linia: 'Kod kres.: <13-cyfrowy EAN>'.
    - W kolejnej niepustej linii powinna być sama liczba (Lp).
    - Potem fragmenty nazwy, aż do wiersza zawierającego 'szt.' i ilość.
    """
    products = []
    for idx, ln in enumerate(all_lines):
        if "Kod kres.:" in ln:
            # EAN
            m_ean = re.search(r"(\d{13})", ln)
            if not m_ean:
                continue
            Barcode_val = m_ean.group(1)

            # Następną linią powinien być Lp (cyfra)
            if idx + 1 < len(all_lines) and all_lines[idx + 1].isdigit():
                Lp_val = int(all_lines[idx + 1])
            else:
                # Jeżeli nie uda się odczytać Lp, pomijamy tę pozycję
                continue

            # Zbieramy fragmenty nazwy od idx+2 aż do wiersza zawierającego 'szt.'
            name_parts = []
            qty = None
            j = idx + 2
            while j < len(all_lines):
                if "szt" in all_lines[j].lower():
                    # np. "96,00 szt."
                    m_qty = re.search(r"([\d\s,]+)\s*szt", all_lines[j], re.IGNORECASE)
                    if m_qty:
                        qty_str = m_qty.group(1).replace(" ", "")
                        qty = float(qty_str.replace(",", "."))
                    break
                else:
                    name_parts.append(all_lines[j])
                j += 1

            Name_val = " ".join(name_parts).strip()
            products.append({
                "Lp": Lp_val,
                "Name": Name_val,
                "Quantity": qty if qty is not None else 0,
                "Barcode": Barcode_val
            })
    return pd.DataFrame(products)


def parse_layout_b(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla układu B – każda pozycja w jednej linii:
      <Lp> <EAN(13)> <pełna nazwa> <ilość>,<xx> szt …
    Przykład:
      1 5029040012366 Nazwa Produktu 96,00 szt.
    """
    products = []
    pattern = re.compile(r"^(\d+)\s+(\d{13})\s+(.+?)\s+([\d,]+)\s+szt", re.IGNORECASE)
    for ln in all_lines:
        m = pattern.match(ln)
        if m:
            Lp_val = int(m.group(1))
            Barcode_val = m.group(2)
            Name_val = m.group(3).strip()
            qty_str = m.group(4).replace(",", ".")
            products.append({
                "Lp": Lp_val,
                "Name": Name_val,
                "Quantity": float(qty_str),
                "Barcode": Barcode_val
            })
    return pd.DataFrame(products)


def parse_layout_c(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla układu C – czysty 13-cyfrowy EAN w osobnej linii, potem Lp, potem nazwa, potem 'szt.' i ilość.
    Przykład:
      5029040012366
      1
      Nazwa Produktu
      96,00 szt.
    """
    products = []
    idx = 0
    while idx < len(all_lines):
        ln = all_lines[idx]
        if re.fullmatch(r"\d{13}", ln):
            Barcode_val = ln
            # Następna linia – Lp
            if idx + 1 < len(all_lines) and all_lines[idx + 1].isdigit():
                Lp_val = int(all_lines[idx + 1])
            else:
                idx += 1
                continue

            # Zbieramy kolejne linie jako nazwę aż trafimy na 'szt.'
            name_parts = []
            qty = None
            j = idx + 2
            while j < len(all_lines):
                if "szt" in all_lines[j].lower():
                    m_qty = re.search(r"([\d\s,]+)\s*szt", all_lines[j], re.IGNORECASE)
                    if m_qty:
                        qty_str = m_qty.group(1).replace(" ", "")
                        qty = float(qty_str.replace(",", "."))
                    break
                else:
                    name_parts.append(all_lines[j])
                j += 1

            Name_val = " ".join(name_parts).strip()
            products.append({
                "Lp": Lp_val,
                "Name": Name_val,
                "Quantity": qty if qty is not None else 0,
                "Barcode": Barcode_val
            })
            # Przeskakujemy indeks do kolejnej pozycji
            idx = j + 1
        else:
            idx += 1

    return pd.DataFrame(products)


# ──────────────────────────────────────────────────────────────────────────────

# 1) FileUploader – wgraj PDF
uploaded_file = st.file_uploader("Wybierz plik PDF", type=["pdf"])
if uploaded_file is None:
    st.stop()

# 2) Odczyt bajtów dokumentu
pdf_bytes = uploaded_file.read()

# 3) Wyciągnięcie linii tekstu (najpierw PyPDF2, potem ewentualnie pdfplumber)
all_lines = extract_text_with_pypdf2(pdf_bytes)

# 4) Jeśli wciąż pusta lista – prawdopodobnie PDF to skan/obraz albo nie udało się nic wydobyć
if not all_lines:
    st.error(
        "Nie udało się odczytać tekstu z PDF-a. "
        "Upewnij się, że to dokument tekstowy, a nie skan. "
        + (
            "\n\nJeśli to „Wydruk.pdf” i masz zainstalowane `pdfplumber`, "
            "dodaj je do środowiska, aby poprawnie wydobyć tekst."
            if not _HAS_PDFPLUMBER
            else ""
        )
    )
    st.stop()

# 5) Wykrycie układu:
#    - Układ B: każda linia zaczyna się od numeru porządkowego i 13-cyfr EAN → regex
#    - Układ C: w all_lines pojawiają się linie, które są dokładnie 13 cyfr → pełna linia to EAN
#    - W przeciwnym razie Układ A
layout = None

pattern_b = re.compile(r"^\d+\s+\d{13}\s+.+\s+[\d,]+\s+szt", re.IGNORECASE)
if any(pattern_b.match(ln) for ln in all_lines):
    layout = "B"
elif any(re.fullmatch(r"\d{13}", ln) for ln in all_lines):
    layout = "C"
else:
    layout = "A"

# 6) Wywołanie odpowiedniego parsera
if layout == "A":
    df = parse_layout_a(all_lines)
elif layout == "B":
    df = parse_layout_b(all_lines)
else:
    df = parse_layout_c(all_lines)

# 7) Jeśli DataFrame jest pusty → wyświetlamy alert
if df.empty:
    st.warning("Nie znaleziono żadnych pozycji w pliku PDF. Sprawdź format pliku.")
    st.stop()

# 8) Wyświetlenie tabeli
st.dataframe(df, use_container_width=True)

# 9) Funkcja konwertująca DataFrame na plik Excel
def convert_df_to_excel(df_in: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_in.to_excel(writer, index=False, sheet_name="Zamówienie")
    return output.getvalue()

# 10) Przycisk do pobrania pliku Excel
excel_data = convert_df_to_excel(df)
st.download_button(
    label="Pobierz wynik jako Excel",
    data=excel_data,
    file_name="parsed_zamowienie.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
