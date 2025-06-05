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
    Parsuje PyPDF2.PdfReader (zamówienie PDF) tak, żeby:
    - poprawnie scalić pozycje rozbite między stronami,
    - odciąć stopki (np. linie zawierające "Strona"),
    - zebrać zawsze 'Kod kres.' nawet gdy wpadnie nad lub pod nagłówkiem "Lp".
    """
    products = []
    current = None
    capture_name = False
    name_lines = []

    for page in reader.pages:
        raw_lines = page.extract_text().split("\n")

        # 1) Odetnij stopkę: wszystko od linii zawierającej "Strona" w dół
        footer_idx = None
        for i, ln in enumerate(raw_lines):
            if "Strona" in ln:  # tu można ewentualnie rozszerzyć o inne słowa stopki
                footer_idx = i
                break

        if footer_idx is not None:
            lines = raw_lines[:footer_idx]
        else:
            lines = raw_lines

        # 2) Znajdź nagłówek "Lp" (jeśli jest) i zapamiętaj jego indeks
        header_idx = None
        for i, ln in enumerate(lines):
            if ln.strip().startswith("Lp"):
                header_idx = i
                break

        # 3) Jeśli nagłówek "Lp" się pojawił, to przed headerem wyciągamy
        #    ewentualne "Kod kres." dla bieżącego current
        if header_idx is not None:
            for ln in lines[:header_idx]:
                stripped = ln.strip()
                if stripped.startswith("Kod kres."):
                    parts = stripped.split(":", maxsplit=1)
                    if len(parts) == 2 and current is not None:
                        barcode = parts[1].strip()
                        if current.get("Barcode") is None:
                            current["Barcode"] = barcode
            start_idx = header_idx + 1
        else:
            start_idx = 0

        # 4) Od start_idx aż do końca (bez stopki) – normalne parsowanie:
        for i in range(start_idx, len(lines)):
            stripped = lines[i].strip()

            # 4a) Jeśli linia zaczyna się od "Kod kres.", przypiszemy do current (jeśli nie ma jeszcze)
            if stripped.startswith("Kod kres."):
                parts = stripped.split(":", maxsplit=1)
                if len(parts) == 2 and current is not None:
                    barcode = parts[1].strip()
                    if current.get("Barcode") is None:
                        current["Barcode"] = barcode
                continue

            # 4b) Jeśli linia to sama liczba (Lp lub Quantity)
            if re.fullmatch(r"\d+", stripped):
                # 4b-i) Jeżeli następna linia to "szt.", to traktujemy tę liczbę jako Quantity
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
                    # 4b-ii) W przeciwnym razie – to nowy Lp → tworzymy current od nowa
                    lp_number = int(stripped)
                    current = {"Lp": lp_number, "Name": None, "Quantity": None, "Barcode": None}
                    products.append(current)
                    capture_name = True
                    name_lines = []
                    continue

            # 4c) Jeśli capture_name==True i niepusta linia → to fragment nazwy produktu
            if capture_name and stripped:
                name_lines.append(stripped)
                continue

            # Pozostałe wiersze (np. ceny, VAT, puste) – ignorujemy

        # Po każdej stronie zachowujemy current, capture_name, name_lines
        # i przechodzimy do następnej – dzięki temu scalimy bloki przerwane między stronami.

    # 5) Po przejściu wszystkich stron – wybieramy tylko te wiersze,
    #    które mają Name i Quantity (reszta to "artefakty").
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

# 3) Parsowanie do DataFrame (z komunikatem w trakcie)
with st.spinner("Analizuję PDF…"):
    df = parse_pdf_to_dataframe(pdf_reader)

# 4) Wyświetlamy wynik w tabeli
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
