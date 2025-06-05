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
    Wgraj plik PDF ze zamówieniem. Aplikacja:
    1) Połączy wszystkie strony w jeden ciąg linii (pomijając stopki i nagłówki „Lp”),
    2) Wyłuska numer pozycji (Lp), fragmenty nazwy (z jednej lub wielu linii),
       ilość (linia „<liczba>” + kolejna linia „szt.”) oraz kod EAN (linia zawierająca „Kod kres.”),
    3) Scal pozycje rozbite na dwie strony (tak, by „Kod kres.” trafił do tej samej pozycji),
    4) Na koniec wyświetli wynik w tabeli oraz da możliwość pobrania pliku Excel.
    """
)


def parse_pdf_to_dataframe(reader: PyPDF2.PdfReader) -> pd.DataFrame:
    """
    Zwraca DataFrame z kolumnami ['Lp', 'Name', 'Quantity', 'Barcode'], wczytując wszystkie
    strony PDF-a i łącząc je w jeden ciąg linii (bez stopek i nagłówków). Dzięki temu kod EAN
    dla produktów rozbitych między stronami trafia poprawnie do ich pozycji.
    """
    products = []
    current = None
    capture_name = False
    name_lines = []

    # 1) Połącz wszystkie strony w jedną listę wierszy (wszystkie linie, bez stopek/nagłówków)
    all_lines = []
    started = False
    for page in reader.pages:
        raw = page.extract_text().split("\n")
        for ln in raw:
            stripped = ln.strip()
            # dopóki nie napotkamy nagłówka "Lp", nic nie zbieramy:
            if not started:
                if stripped.startswith("Lp") and not stripped.isdigit():
                    started = True
                continue
            # po wykryciu "Lp" – dodawajemy linie aż do napotkania stopki ("Strona")
            if "Strona" in stripped:
                break
            # pomijamy każdą kolejną linię nagłówka "Lp" (jeśli się pojawi ponownie)
            if stripped.startswith("Lp") and not stripped.isdigit():
                continue
            # w pozostałych wierszach zapisujemy;text
            all_lines.append(stripped)

    # 2) Iterujemy po all_lines i wyłuskujemy Lp, Name, Quantity oraz Barcode:
    i = 0
    while i < len(all_lines):
        ln = all_lines[i]

        # 2a) Jeżeli wiersz zawiera "Kod kres", wyciągamy po ":" kod EAN i przypisujemy do
        #     bieżącej pozycji (jeśli jeszcze nie ma Barcode):
        if "Kod kres" in ln:
            parts = ln.split(":", maxsplit=1)
            if len(parts) == 2 and current:
                ean = parts[1].strip()
                if not current.get("Barcode"):
                    current["Barcode"] = ean
            i += 1
            continue

        # 2b) Jeżeli ln to sama liczba – może to być Lp albo Quantity:
        if re.fullmatch(r"\d+", ln):
            nxt = all_lines[i + 1] if (i + 1) < len(all_lines) else ""

            # 2b-i) Jeśli następna linia to "szt.", to traktujemy ln jako Quantity:
            if nxt.lower() == "szt.":
                qty = int(ln)
                if current:
                    current["Quantity"] = qty
                    full_name = " ".join(name_lines).strip()
                    current["Name"] = full_name
                    name_lines = []
                    capture_name = False
                i += 2  # pomijamy także linię "szt."
                continue

            # 2b-ii) Jeśli następna linia zawiera litery (a nie zaczyna się od "Kod kres"), 
            #        uznajemy ln za nowy numer pozycji (Lp):
            elif re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", nxt) and not nxt.startswith("Kod kres"):
                Lp = int(ln)
                current = {"Lp": Lp, "Name": None, "Quantity": None, "Barcode": None}
                products.append(current)
                capture_name = True
                name_lines = []
                i += 1
                continue

            # 2b-iii) W przeciwnym razie to „liczbowe artefakty” (np. cena czy numer w stopce) – pomijamy:
            else:
                i += 1
                continue

        # 2c) Jeżeli capture_name == True i ln nie jest pusty → to fragment nazwy:
        if capture_name:
            #  - Linie w formacie „123,45” lub „1 234,56” to ceny – pomiń:
            if re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", ln):
                i += 1
                continue
            #  - Linie zaczynające się od "VAT" to nagłówki VAT – pomiń:
            if ln.startswith("VAT"):
                i += 1
                continue
            #  - Jednoznacznie "/" to osobny wiersz opisu (pomijamy):
            if ln == "/":
                i += 1
                continue
            #  - Jeśli linia zawiera przynajmniej jedną literę (polską albo łacińską), traktujemy ją jako część nazwy:
            if re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", ln):
                name_lines.append(ln)
                i += 1
                continue
            #  - W przeciwnym razie nie jest to ani nazwa, ani cena, ani VAT – pomijamy:
            i += 1
            continue

        # 2d) W innym wypadku (puste wiersze itp.) – pomiń:
        i += 1

    # 3) Na koniec budujemy DataFrame i odrzucamy wiersze niekompletne (brak nazwy lub ilości)
    df = pd.DataFrame(products)
    df = df.dropna(subset=["Name", "Quantity"]).reset_index(drop=True)
    return df


# ──────────────────────────────────────────────────────────────────────────────

# 1) File Uploader: wczytanie PDF-a
uploaded_file = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if uploaded_file is None:
    st.info("Proszę wgrać plik PDF, aby uruchomić parser.")
    st.stop()

# 2) Odczyt PDF-a
try:
    pdf_reader = PyPDF2.PdfReader(uploaded_file)
except Exception as e:
    st.error(f"Nie udało się wczytać PDF-a: {e}")
    st.stop()

# 3) Parsowanie do DataFrame
with st.spinner("Łączę strony i analizuję PDF…"):
    df = parse_pdf_to_dataframe(pdf_reader)

# 4) Wyświetlenie wyniku w tabeli
st.subheader("Wyekstrahowane pozycje zamówienia")
st.dataframe(df, use_container_width=True)

# 5) Pobranie wyniku jako plik Excel
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
