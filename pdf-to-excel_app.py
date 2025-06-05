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
    Wgraj plik PDF ze zamówieniem. Skrypt połączy wszystkie strony w jedną listę wierszy 
    (pomiatając stopki i nagłówki „Lp”), a następnie wyciągnie:
    - numer pozycji (Lp),
    - wszystkie fragmenty nazwy (z jednej lub wielu linii),
    - ilość (linia „<liczba>” + kolejna linia „szt.”),
    - kod EAN (linia zawierająca „Kod kres.”).

    Dzięki temu pozycje rozbite na dwie strony (gdzie „Kod kres.” trafia dopiero później) 
    zostaną poprawnie scalone i mają przypisany EAN.
    """
)


def parse_pdf_to_dataframe(reader: PyPDF2.PdfReader) -> pd.DataFrame:
    """
    1. Łączy wszystkie strony PDF w jeden ciąg linii, pomijając:
       - każdą linię od słowa "Strona" w dół (stopka),
       - każdą linię rozpoczynającą się od "Lp" niebędącą liczbową (nagłówek tabeli).

    2. Przechodzi kolejno po tej zunifikowanej liście:
       - Gdy znajdzie "Kod kres" → wyciąga kod i przypisuje go do bieżącej pozycji.
       - Gdy linia to sama liczba, a następna to "szt." → to jest Quantity.
       - Gdy linia to sama liczba, a następna to nie "szt." → to jest nowy Lp.
       - W trybie capture_name (po wykryciu nowego Lp) każda niepusta linia 
         (która nie jest czystą liczbą ani "Kod kres") dokłada się do name_lines.
    3. Po zakończeniu pierwszego przejścia buduje DataFrame i odrzuca wiersze 
       bez nazwy lub ilości.
    """
    products = []
    current = None
    capture_name = False
    name_lines = []

    # 1) Połącz wszystkie strony w jedną listę linii, pomijając stopki i nagłówki "Lp"
    all_lines = []
    for page in reader.pages:
        raw = page.extract_text().split("\n")
        for ln in raw:
            stripped = ln.strip()
            # jeśli napotkamy stopkę (linia zawiera "Strona"), przerywamy tę stronę
            if "Strona" in stripped:
                break
            # jeśli linia to nagłówek tabeli, tzn. zaczyna się od "Lp" i nie jest czystą liczbą, pomijamy
            if stripped.startswith("Lp") and not stripped.isdigit():
                continue
            # w każdym innym wypadku dopisujemy do all_lines (może to być fragment nazwy, ilość, kod, itp.)
            all_lines.append(stripped)

    # 2) Przejdź kolejno po all_lines
    i = 0
    while i < len(all_lines):
        ln = all_lines[i]

        # 2a) Jeśli linia zawiera "Kod kres", wyciągnij kod po ":" i przypisz do current
        if "Kod kres" in ln:
            parts = ln.split(":", maxsplit=1)
            if len(parts) == 2 and current is not None:
                ean = parts[1].strip()
                if not current.get("Barcode"):
                    current["Barcode"] = ean
            i += 1
            continue

        # 2b) Jeśli ln to sama liczba → może być Lp lub Quantity
        if re.fullmatch(r"\d+", ln):
            # 2b-i) Jeśli następna linia to "szt.", traktujemy ln jako Quantity
            if i + 1 < len(all_lines) and all_lines[i + 1].lower() == "szt.":
                qty = int(ln)
                if current is not None:
                    current["Quantity"] = qty
                    # po ustaleniu ilości łączymy name_lines w jedno pole Name
                    full_name = " ".join(name_lines).strip()
                    current["Name"] = full_name
                    name_lines = []
                    capture_name = False
                i += 2  # pomijamy też wiersz "szt."
                continue
            else:
                # 2b-ii) To jest nowy Lp → zakładamy początek nowego produktu
                Lp = int(ln)
                current = {"Lp": Lp, "Name": None, "Quantity": None, "Barcode": None}
                products.append(current)
                capture_name = True
                name_lines = []
                i += 1
                continue

        # 2c) Jeśli capture_name=True i ln nie jest pusty → to fragment nazwy
        if capture_name and ln:
            name_lines.append(ln)
            i += 1
            continue

        # 2d) Wszystkie inne wiersze (np. puste, zawierające cenę/VAT itp.) pomijamy
        i += 1

    # 3) Zbuduj DataFrame i odfiltruj wiersze niekompletne
    df = pd.DataFrame(products)
    df = df.dropna(subset=["Name", "Quantity"]).reset_index(drop=True)
    return df


# ──────────────────────────────────────────────────────────────────────────────

# FileUploader: pozwala wgrać PDF
uploaded_file = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if uploaded_file is None:
    st.info("Proszę wgrać plik PDF, aby uruchomić parser.")
    st.stop()

# Czytanie PDF-a
try:
    pdf_reader = PyPDF2.PdfReader(uploaded_file)
except Exception as e:
    st.error(f"Nie udało się wczytać PDF-a: {e}")
    st.stop()

# Parsowanie do DataFrame
with st.spinner("Łączę strony i analizuję PDF…"):
    df = parse_pdf_to_dataframe(pdf_reader)

# Wyświetlenie tabeli w Streamlit
st.subheader("Wyekstrahowane pozycje zamówienia")
st.dataframe(df, use_container_width=True)

# Przygotowanie przycisku do pobrania jako Excel
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
