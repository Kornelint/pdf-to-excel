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
    1. Połączy wszystkie strony w jeden ciąg wierszy, pomijając:
       - stopki (linie zaczynające się od "Strona", "Wydrukowano", lub zawierające "ZD <…>"),
       - pierwsze powtarzające się nagłówki tabel (np. linie zaczynające się od "Lp" nie będące czystą liczbą, 
         czy "Nazwa towaru lub usługa", „Ilość”, „J. miary”, „Cena”, „Wartość”, „netto”, „brutto”).  
    2. Dla każdej linii:
       - Jeśli linia to sama liczba a następna linia zawiera litery → to jest **Lp** (numer pozycji).
       - Gdy `capture_name=True`, każda linia zawierająca litery (nie wyglądająca na cenę ani nagłówek VAT) 
         dokłada się do fragmentów nazwy (`Name`).
       - Jeśli napotkamy wiersz będący liczbą, a wiersz dalej to `"szt."` → to jest **Quantity**.
         Po przypisaniu ilości łączymy wszystkich zebrane fragmenty nazwy w jeden ciąg.
       - Jeśli linia zawiera frazę `"Kod kres."` → wyciągamy kod EAN i przypisujemy do bieżącej pozycji.
    3. Na koniec wyświetli tabelę ze wszystkimi produktami i pozwoli pobrać plik Excel 
       z kolumnami: `Lp`, `Name`, `Quantity`, `Barcode`.
    """
)

def parse_pdf_generic(reader: PyPDF2.PdfReader) -> pd.DataFrame:
    """
    Uniwersalny parser, który poradzi sobie zarówno z formatem, gdzie w nagłówku jest "Lp.",
    jak i z formatem, gdzie nagłówek to "Nazwa towaru lub usługa". 
    Wszystkie strony łączymy w all_lines, usuwając stopki i powtórzone nagłówki, 
    a potem przechodzimy kolejno po all_lines, wyłuskując:
      - Lp (pierwsza czysta liczba, po której następuje linia z literami),
      - Name (wszystkie kolejne wiersze z literami, aż do linii z ilością),
      - Quantity (linia = liczba, a następna linia to "szt."),
      - Barcode (linia zawierająca "Kod kres.: <EAN>").
    """
    products = []
    current = None
    capture_name = False
    name_fragments = []

    all_lines = []
    started = False

    # 1) Zbierz wszystkie strony w all_lines, pomijając stopki i początkowe nagłówki
    for page in reader.pages:
        raw = page.extract_text().split("\n")
        for ln in raw:
            stripped = ln.strip()

            # 1a) Jeśli napotkamy stopkę ("Wydrukowano", "Strona", lub "ZD <cyfry>"), przerwij tę stronę:
            if stripped.startswith("Wydrukowano") or stripped.startswith("Strona") or re.match(r"ZD \d+/?\d*", stripped):
                break

            # 1b) Dopóki nie zaczniemy parsować danych, pomijamy wiersze wstępu:
            if not started:
                # Jeśli natrafimy na linię rozpoczynającą się od "Lp" (nie będącą czystą liczbą) → format z nagłówkiem Lp
                # Lub na linię zaczynającą się od "Nazwa towaru" → format z nagłówkiem nazwy towaru
                if (stripped.startswith("Lp") and not stripped.isdigit()) or stripped.lower().startswith("nazwa towaru"):
                    started = True
                    # nie dodajemy tej linii (nagłówek)
                    continue
                # Jeśli nie ma jeszcze nagłówka, pomijamy
                continue

            # 1c) Kiedy już started=True, pomijamy typowe wiersze nagłówków w tabeli, np.:
            if stripped.startswith("Lp") and not stripped.isdigit():
                continue
            if stripped.lower().startswith("nazwa towaru"):
                continue
            if stripped in ["Ilość", "J. miary", "Cena", "Wartość", "netto", "brutto", "Indeks katalogowy", ""]:
                continue

            # 1d) W innych przypadkach dopisujemy stripped do all_lines
            all_lines.append(stripped)

    # 2) Przechodzimy po all_lines sekwencyjnie
    i = 0
    n = len(all_lines)
    while i < n:
        ln = all_lines[i]

        # 2a) Jeżeli linia zawiera "Kod kres", wyciągamy EAN i przypisujemy do bieżącej pozycji
        if "Kod kres" in ln:
            parts = ln.split(":", maxsplit=1)
            if len(parts) == 2 and current:
                ean = parts[1].strip()
                if not current.get("Barcode"):
                    current["Barcode"] = ean
            i += 1
            continue

        # 2b) Jeżeli ln to czysta liczba → może to być Lp albo Quantity
        if re.fullmatch(r"\d+", ln):
            nxt = all_lines[i + 1] if (i + 1) < n else ""

            # 2b-i) Jeżeli następna linia to dokładnie "szt." → ln to Quantity
            if nxt.lower() == "szt.":
                qty = int(ln)
                if current:
                    current["Quantity"] = qty
                    # Po wpisaniu Quantity łączymy wszystkie zebrane fragmenty nazwy:
                    current["Name"] = " ".join(name_fragments).strip()
                    name_fragments = []
                    capture_name = False
                i += 2  # przeskocz także "szt."
                continue

            # 2b-ii) W przeciwnym razie jeśli następna linia zawiera litery (i nie zaczyna się od "Kod kres") → ln to nowy Lp
            elif re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", nxt) and not nxt.startswith("Kod kres"):
                Lp_val = int(ln)
                current = {"Lp": Lp_val, "Name": None, "Quantity": None, "Barcode": None}
                products.append(current)
                capture_name = True
                name_fragments = []
                i += 1
                continue

            # 2b-iii) W przeciwnym razie – prawdopodobnie cyfra to fragment ceny lub inny „artefakt” – pomiń
            else:
                i += 1
                continue

        # 2c) Jeżeli capture_name=True i ln zawiera litery → dodajemy do name_fragments
        if capture_name and re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", ln):
            # Pomijamy linie wyglądające jak ceny np. "123,45" lub "1 234,56":
            if re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", ln):
                i += 1
                continue
            # Pomijamy linie zaczynające się od "VAT"
            if ln.startswith("VAT"):
                i += 1
                continue
            # Pomijamy pojedyncze "/"
            if ln == "/":
                i += 1
                continue

            # W pozostałych przypadkach traktujemy ln jako fragment nazwy:
            name_fragments.append(ln)
            i += 1
            continue

        # 2d) W innym wypadku (puste linie, nagłówki kolumn, wartości cenowe itp.) – po prostu pomiń
        i += 1

    # 3) Po zakończeniu pętli budujemy DataFrame i odrzucamy wiersze bez Name lub bez Quantity
    df = pd.DataFrame(products)
    df = df.dropna(subset=["Name", "Quantity"]).reset_index(drop=True)
    return df


# ──────────────────────────────────────────────────────────────────────────────

# FileUploader: wgraj PDF
uploaded_file = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if uploaded_file is None:
    st.info("Proszę wgrać plik PDF, aby uruchomić parser.")
    st.stop()

# Spróbuj wczytać PDF przez PyPDF2
try:
    pdf_reader = PyPDF2.PdfReader(uploaded_file)
except Exception as e:
    st.error(f"Nie udało się wczytać PDF-a: {e}")
    st.stop()

# Parsowanie do DataFrame (z komunikatem spinnera)
with st.spinner("Łączę strony i analizuję PDF…"):
    df = parse_pdf_generic(pdf_reader)

# Wyświetlenie rezultatu w tabeli
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
