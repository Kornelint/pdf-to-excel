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

    Zasady:
    1. Zbiera wszystkie linie (bez stopki "Strona") w kolejności, jak w PDF-ie.
    2. Pomija linie nagłówka tabeli zaczynające się od "Lp" (ale nie czyste liczby).
    3. Jeśli linia zawiera frazę "Kod kres", wyciąga kod i przypisuje go do bieżącej pozycji,
       o ile nie ma jeszcze przypisanego barcode.
    4. Jeśli linia to sama liczba, a następna linia to "szt.", traktuje to jako Quantity.
    5. Jeśli linia to sama liczba i następna linia ≠ "szt.", to jest nowy Lp.
    6. Jeżeli capture_name=True i linia nie-pusta, traktuje ją jako fragment nazwy.
    7. Na końcu usuwa wiersze, które nie mają ani Name, ani Quantity.
    """
    products = []
    current = None
    capture_name = False
    name_lines = []

    # 1) Zbierz wszystkie linie z każdej strony, pomijając stopki i nagłówki “Lp”
    all_lines = []
    for page in reader.pages:
        raw = page.extract_text().split("\n")
        for ln in raw:
            stripped = ln.strip()
            # Jeśli spotkamy linię stopki (zawiera "Strona"), przerywamy tę stronę
            if "Strona" in stripped:
                break
            # Jeśli linia to nagłówek tabeli zaczynający się od "Lp" i nie jest czystą liczbą – pomiń
            if stripped.startswith("Lp") and not stripped.isdigit():
                continue
            all_lines.append(stripped)

    # 2) Iterujemy po wszystkich liniach w kolejności
    i = 0
    while i < len(all_lines):
        ln = all_lines[i]

        # 2a) Jeżeli linia zawiera "Kod kres", to wyciągnij EAN i przypisz do bieżącej pozycji
        if "Kod kres" in ln:
            parts = ln.split(":", maxsplit=1)
            if len(parts) == 2 and current is not None:
                ean = parts[1].strip()
                if not current.get("Barcode"):
                    current["Barcode"] = ean
            i += 1
            continue

        # 2b) Czy ln to sama liczba? (może to być Lp lub Quantity)
        if re.fullmatch(r"\d+", ln):
            # 2b-i) Jeżeli następna linia to "szt.", to to jest Quantity
            if i + 1 < len(all_lines) and all_lines[i + 1].lower() == "szt.":
                qty = int(ln)
                if current is not None:
                    current["Quantity"] = qty
                    # Po uzupełnieniu ilości – łączymy nazwy
                    full_name = " ".join(name_lines).strip()
                    current["Name"] = full_name
                    name_lines = []
                    capture_name = False
                i += 2  # pomijamy także "szt."
                continue
            else:
                # 2b-ii) To nowy Lp
                Lp = int(ln)
                current = {"Lp": Lp, "Name": None, "Quantity": None, "Barcode": None}
                products.append(current)
                capture_name = True
                name_lines = []
                i += 1
                continue

        # 2c) Jeżeli capture_name=True i ln nie-puste → fragment nazwy
        if capture_name and ln:
            name_lines.append(ln)
            i += 1
            continue

        # 2d) Wszelkie inne wiersze (np. ceny, VAT, puste) – pomijamy
        i += 1

    # 3) Po przetworzeniu wszystkich linii, utwórz DataFrame i usuń wiersze niekompletne
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
