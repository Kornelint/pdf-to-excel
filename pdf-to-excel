# app.py

import streamlit as st
import pandas as pd
import re
import PyPDF2
import io

st.set_page_config(page_title="PDF → Excel", layout="wide")
st.title("Konwerter zamówienia PDF → Excel")

st.markdown(
    """
    Ten prosty demo-aplikacja Streamlit pozwala wczytać plik PDF 
    (zamówienie, gdzie każda pozycja rozbita jest na kilka wierszy) i 
    wyeksportować go do Excela.  
    Plik musi zawierać nagłówek „Lp” oraz wiersze w układzie:  
    - numer pozycji (liczba)  
    - fragmenty nazwy (może być 2–3 wiersze)  
    - ilość (liczba) i wiersz „szt.”  
    - na końcu blok z „Kod kres.: XXXXXXXX”  
    """
)

# 1) Użytkownik wrzuca PDF
uploaded_file = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if uploaded_file is None:
    st.info("Proszę wgrać plik PDF, aby uruchomić parser.")
    st.stop()

# 2) Gdy użytkownik wrzuci plik, odczytujemy jego zawartość
try:
    # PyPDF2 oczekuje pliku "readable()" → dlatego wyciągamy bajty
    pdf_reader = PyPDF2.PdfReader(uploaded_file)
except Exception as e:
    st.error(f"Nie udało się wczytać PDF-a: {e}")
    st.stop()

# 3) Funkcja, która na podstawie wczytanych stron tworzy DataFrame
def parse_pdf_to_dataframe(reader: PyPDF2.PdfReader) -> pd.DataFrame:
    products = []
    current = None
    capture_name = False
    name_lines = []

    # Iterujemy po stronach
    for page in reader.pages:
        # Podziel tekst na linie
        raw_lines = page.extract_text().split("\n")

        # 1. Odetnij stopkę (np. „Strona X”)
        footer_idx = None
        for i, ln in enumerate(raw_lines):
            if "Strona" in ln:   # lub inny słownik stopki w Twoich PDF-ach
                footer_idx = i
                break

        if footer_idx is not None:
            lines = raw_lines[:footer_idx]
        else:
            lines = raw_lines

        # 2. Sprawdź, czy jest nagłówek „Lp” na tej stronie
        if any(line.strip().startswith("Lp") for line in lines):
            header_idx = next(i for i, line in enumerate(lines) if line.strip().startswith("Lp"))
            start_idx = header_idx + 1
        else:
            start_idx = 0

        # 3. Parsowanie linii od start_idx do końca (albo do stopki)
        for i in range(start_idx, len(lines)):
            stripped = lines[i].strip()

            # 3a) Kod kreskowy
            if stripped.startswith("Kod kres."):
                parts = stripped.split(":", maxsplit=1)
                if len(parts) == 2 and current is not None:
                    barcode = parts[1].strip()
                    if current.get("Barcode") is None:
                        current["Barcode"] = barcode
                continue

            # 3b) Czy to liczba? (może to być Lp lub Quantity)
            if re.fullmatch(r"\d+", stripped):
                # 3b-i) Jeśli następny wiersz to "szt." → Quantity
                if i + 1 < len(lines) and lines[i + 1].strip().lower() == "szt.":
                    qty = int(stripped)
                    if current is not None:
                        current["Quantity"] = qty
                        # Sklej fragmenty nazwy
                        full_name = " ".join(name_lines).strip()
                        current["Name"] = full_name
                        name_lines = []
                        capture_name = False
                    continue
                else:
                    # 3b-ii) To nowy Lp → zaczynamy nową pozycję
                    lp_number = int(stripped)
                    current = {"Lp": lp_number, "Name": None, "Quantity": None, "Barcode": None}
                    products.append(current)
                    capture_name = True
                    name_lines = []
                    continue

            # 3c) Jeżeli zbieramy nazwę (capture_name==True) i wiersz nie jest pusty
            if capture_name and stripped:
                name_lines.append(stripped)
                continue

            # Pozostałe wiersze pomijamy

    # 4) Po przejrzeniu wszystkich stron filtrujemy puste/niekompletne wpisy
    df = pd.DataFrame(products)
    df = df.dropna(subset=["Name", "Quantity"]).reset_index(drop=True)
    return df

# Wywołujemy parser
with st.spinner("Analizuję PDF…"):
    df = parse_pdf_to_dataframe(pdf_reader)

# 5) Wyświetlamy tabelę w Streamlit
st.subheader("Wyekstrahowane pozycje zamówienia")
st.dataframe(df)  # pozwala scrollować wiersze/kolumny

# 6) Dajmy opcję pobrania wyniku jako Excel
def convert_df_to_excel(df_in: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    # Używamy pandas i openpyxl, żeby zapisać w pamięci plik xlsx
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_in.to_excel(writer, index=False, sheet_name="Zamowienie")
    data = output.getvalue()
    return data

excel_data = convert_df_to_excel(df)
st.download_button(
    label="Pobierz wynik jako Excel",
    data=excel_data,
    file_name="parsed_zamowienie.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
