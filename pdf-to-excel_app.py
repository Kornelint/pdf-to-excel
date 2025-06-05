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
    Wgraj plik PDF ze zamówieniem, a ja wyciągnę wszystkie pozycje (nawet te rozbite między stronami)
    i udostępnię wynik jako tabelę oraz plik Excel.
    """
)


def parse_pdf_to_dataframe(reader: PyPDF2.PdfReader) -> pd.DataFrame:
    """
    Parsuje PDF zamówienia w formacie PyPDF2.PdfReader i zwraca DataFrame z kolumnami:
    ['Lp', 'Name', 'Quantity', 'Barcode'].

    - Odcina stopki (linie zawierające "Strona").
    - Łapie wszystkie wystąpienia "Kod kres" w całym bloku tekstu, przypisując je do bieżącej pozycji.
    - Scalanie bloków produktów rozbitych między stronami.
    - Na końcu usuwa wiersze, które nie mają nazwy lub ilości.
    """
    products = []
    current = None
    capture_name = False
    name_lines = []

    for page in reader.pages:
        raw_lines = page.extract_text().split("\n")

        # 1) Odetnij stopkę: wszystko od momentu, gdy pojawi się wiersz zawierający "Strona"
        footer_idx = None
        for i, ln in enumerate(raw_lines):
            if "Strona" in ln:
                footer_idx = i
                break

        if footer_idx is not None:
            lines = raw_lines[:footer_idx]
        else:
            lines = raw_lines

        # 2) Sprawdź, czy jest nagłówek "Lp" i zapamiętaj jego indeks
        header_idx = None
        for i, ln in enumerate(lines):
            if ln.strip().startswith("Lp"):
                header_idx = i
                break

        # 3) W całym bloku 'lines' znajdź wszystkie "Kod kres" i przypisz je do bieżącej pozycji
        for ln in lines:
            stripped = ln.strip()
            if "Kod kres" in stripped:
                parts = stripped.split(":", maxsplit=1)
                if len(parts) == 2 and current is not None:
                    candidate = parts[1].strip()
                    if not current.get("Barcode"):  # przypisz, jeżeli jest puste
                        current["Barcode"] = candidate

        # 4) Ustal od którego wiersza (start_idx) zaczynamy parsować tabelę:
        #    - jeśli header_idx istnieje, to start_idx = header_idx + 1
        #    - jeśli nie, to from 0 (kontynuacja rozbitego bloku)
        if header_idx is not None:
            start_idx = header_idx + 1
        else:
            start_idx = 0

        # 5) Od start_idx do końca 'lines' – normalne parsowanie Lp → nazwa → ilość:
        for i in range(start_idx, len(lines)):
            stripped = lines[i].strip()

            # 5a) Jeśli linia zawiera "Kod kres", przeskoczemy, bo już to złapaliśmy w pętli wyżej
            if "Kod kres" in stripped:
                continue

            # 5b) Jeżeli to sama liczba (może być Lp lub Quantity)
            if re.fullmatch(r"\d+", stripped):
                # 5b-i) Jeżeli następna linia to "szt.", traktujemy tę liczbę jako Quantity
                if i + 1 < len(lines) and lines[i + 1].strip().lower() == "szt.":
                    qty = int(stripped)
                    if current is not None:
                        current["Quantity"] = qty
                        full_name = " ".join(name_lines).strip()
                        current["Name"] = full_name
                        name_lines = []
                        capture_name = False
                    continue
                else:
                    # 5b-ii) W przeciwnym razie to nowy Lp → tworzymy nowy słownik
                    lp_number = int(stripped)
                    current = {"Lp": lp_number, "Name": None, "Quantity": None, "Barcode": None}
                    products.append(current)
                    capture_name = True
                    name_lines = []
                    continue

            # 5c) Jeśli capture_name=True i linia nie jest pusta → fragment nazwy produktu
            if capture_name and stripped:
                name_lines.append(stripped)
                continue

            # Pozostałe wiersze (np. ceny, VAT, puste) ignorujemy

        # Koniec przetwarzania tej strony → przejdź dalej, zachowując bieżący 'current'

    # 6) Po przejściu wszystkich stron: stwórz DataFrame i odrzuć niekompletne wiersze
    df = pd.DataFrame(products)
    df = df.dropna(subset=["Name", "Quantity"]).reset_index(drop=True)
    return df


# ──────────────────────────────────────────────────────────────────────────────

# 1) FileUploader: wczytanie PDF-a
uploaded_file = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if uploaded_file is None:
    st.info("Proszę wgrać plik PDF, aby uruchomić parser.")
    st.stop()

# 2) Czytanie PDF-a
try:
    pdf_reader = PyPDF2.PdfReader(uploaded_file)
except Exception as e:
    st.error(f"Nie udało się wczytać PDF-a: {e}")
    st.stop()

# 3) Parsowanie do DataFrame
with st.spinner("Analizuję PDF…"):
    df = parse_pdf_to_dataframe(pdf_reader)

# 4) Wyświetlenie wyniku w tabeli
st.subheader("Wyekstrahowane pozycje zamówienia")
st.dataframe(df, use_container_width=True)

# 5) Opcja pobrania jako plik Excel
def convert_df_to_excel(df_in: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_in.to_excel(writer, index=False, sheet_name="Zamowienie")
    return output.getvalue()

excel_data = convert_df_to_excel(df)
st.download_button(
    label="Pobierz wynik jako Excel",
    data=excel_data,
    file_name="parsed_zamowienie.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
