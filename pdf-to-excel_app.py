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
    1. Przechodzimy po kolejnych liniach wszystkich stron (bez stopki "Strona").
    2. Jeśli linia zawiera wyłącznie liczbę, a następna linia to "szt.", to to jest ilość (Quantity).
    3. Jeśli linia zawiera wyłącznie liczbę, a następna linia to nie "szt.", to to jest nowy Lp.
    4. Jeśli linia zawiera frazę "Kod kres", to wyciągamy po niej kod i przypisujemy do bieżącej pozycji.
    5. Wszystkie inne linie (fragmenty nazwy lub ceny czy VAT) traktujemy tak:
       - Jeśli capture_name=True, a linia nie jest pusta i nie jest liczba, dopisujemy ją jako fragment nazwy.
       - Jeśli linia jest pusta lub zawiera cenę czy VAT, pozycja nazwy się domyka dopiero, gdy natrafimy na ilość.
    6. Na końcu odrzucamy wiersze bez Name lub bez Quantity.
    """
    products = []
    current = None
    capture_name = False
    name_lines = []

    # Zbierz wszystkie linie (bez stopki "Strona …")
    all_lines = []
    for page in reader.pages:
        raw = page.extract_text().split("\n")
        for ln in raw:
            if "Strona" in ln:
                break  # ignorujemy wszystko od "Strona" dalej na tej stronie
            all_lines.append(ln)

    # Przechodzimy po wszystkich liniach w kolejności, jak w PDF-ie
    i = 0
    while i < len(all_lines):
        ln = all_lines[i].strip()

        # 1) "Kod kres" w dowolnym miejscu linii?
        if "Kod kres" in ln:
            parts = ln.split(":", maxsplit=1)
            if len(parts) == 2 and current is not None:
                ean = parts[1].strip()
                if not current.get("Barcode"):
                    current["Barcode"] = ean
            i += 1
            continue

        # 2) Czy linia to sama liczba?
        if re.fullmatch(r"\d+", ln):
            # 2a) Jeżeli następna linia to "szt.", to traktujemy ln jako Quantity
            if i + 1 < len(all_lines) and all_lines[i + 1].strip().lower() == "szt.":
                qty = int(ln)
                if current is not None:
                    current["Quantity"] = qty
                    # Po ilości znaczy: nazwa się kończy, scal fragmenty
                    full_name = " ".join(name_lines).strip()
                    current["Name"] = full_name
                    name_lines = []
                    capture_name = False
                i += 2  # pomijamy także "szt."
                continue
            else:
                # 2b) W przeciwnym wypadku ln to nowy Lp
                Lp = int(ln)
                current = {"Lp": Lp, "Name": None, "Quantity": None, "Barcode": None}
                products.append(current)
                capture_name = True
                name_lines = []
                i += 1
                continue

        # 3) Jeżeli capture_name=True i linia nie jest pusta → fragment nazwy
        if capture_name and ln:
            name_lines.append(ln)
            i += 1
            continue

        # 4) Wszelkie inne wiersze (np. puste, ceny, VAT) po prostu pomijamy
        i += 1

    # 5) Po przetworzeniu wszystkich linii, utwórz DataFrame i przefiltruj
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
