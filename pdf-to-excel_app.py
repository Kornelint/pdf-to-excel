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
    1) Połączy wszystkie strony w jeden ciąg wierszy (pomijając stopki i powtórzone nagłówki „Lp”),
    2) Wyłuska numer pozycji (Lp), wszystkie fragmenty nazwy, ilość („<liczba>” + „szt.”)
       oraz kod EAN (linia zawierająca „Kod kres.”),
    3) Scal pozycje rozbite na dwie strony (tak, aby kod EAN zawsze trafił do tej samej pozycji),
    4) Wyświetli wynik w tabeli oraz umożliwi pobranie go w postaci pliku Excel.
    """
)


def parse_pdf_to_dataframe(reader: PyPDF2.PdfReader) -> pd.DataFrame:
    """
    Parsuje PDF zamówienia w formacie PyPDF2.PdfReader i zwraca DataFrame z kolumnami:
    ['Lp', 'Name', 'Quantity', 'Barcode'].

    Algorytm:
    1. Łączy wszystkie strony w jeden ciąg „all_lines”, pomijając:
       - każdą linię od momentu napotkania słowa "Wydrukowano" lub nagłówka numeru dokumentu ("ZD ...") albo "Strona" (stopka),
       - każdą linię nagłówka tabeli zaczynającą się od "Lp" i niebędącą czystą liczbą,
       - wszystkie linie aż do pierwszego prawdziwego numeru pozycji (Lp), by pominąć inne teksty nagłówka PDF-a.
    2. Przechodzi po all_lines sekwencyjnie:
       - Gdy wiersz zawiera frazę "Kod kres", wyciąga to, co po „:”, i jeśli bieżąca pozycja (`current`) nie ma jeszcze 
         kodu (`Barcode`), przypisuje mu wyciągnięty ciąg cyfr.
       - Gdy wiersz to sama liczba i kolejna linia to "szt.", traktuje tę liczbę jako `Quantity` (ilość). Po przypisaniu 
         łączy wszystkie zebrane fragmenty nazwy w jedno pole `Name` i resetuje `capture_name`.
       - Gdy wiersz to sama liczba i kolejna linia zawiera litery (nazwę), traktuje tę liczbę jako nowy numer pozycji (`Lp`),
         tworzy nowy słownik `current = {'Lp': <liczba>, 'Name': None, 'Quantity': None, 'Barcode': None}` i włącza 
         `capture_name = True`, aby kolejne nie-puste linie uznać za fragmenty nazwy.
       - Jeśli `capture_name == True` i bieżący wiersz nie jest liczbą ani "szt." ani fragmentem kodu, traktuje go jako 
         kolejny fragment nazwy (`name_lines.append(ln)`).
       - Wszystkie inne wiersze (np. ceny, VAT, puste linie itp.) są pomijane.
    3. Po przejściu wszystkich wierszy tworzy pandas.DataFrame z listy `products` i odfiltrowuje wiersze,
       które nie mają jednocześnie `Name` i `Quantity` (bo to zwykle „szum” z parsowania).
    """
    products = []
    current = None
    capture_name = False
    name_lines = []

    all_lines = []
    started = False

    # 1) Połącz wszystkie strony w jeden ciąg wierszy, omijając stopki i nagłówki
    for page in reader.pages:
        raw = page.extract_text().split("\n")
        for ln in raw:
            stripped = ln.strip()

            # Jeśli to stopka (zawiera "Wydrukowano", numer dokumentu "ZD ..." lub "Strona"), przerywamy tę stronę:
            if stripped.startswith("Wydrukowano") or re.match(r"ZD \d{4}", stripped) or stripped.startswith("Strona"):
                break

            if not started:
                # Początek parsowania: szukamy pierwszego prawdziwego numeru pozycji
                # Jeżeli napotkamy liczbę, a kolejna linia zawiera litery → to jest Lp
                if re.fullmatch(r"\d+", stripped):
                    # Znajdź następny wiersz w „raw”
                    idx_in_raw = raw.index(ln)
                    nxt = raw[idx_in_raw + 1].strip() if (idx_in_raw + 1) < len(raw) else ""
                    if re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", nxt):
                        # Od tego momentu zaczynamy zbierać rzeczywiste dane zamówienia
                        started = True
                        all_lines.append(stripped)
                        continue
                # dopóki nie napotkamy prawdziwej Lp, przeskakujemy wiersze
                continue

            # Jeżeli już zaczęliśmy, pomijamy powtórzone nagłówki tabeli (np. linia zaczynająca się od "Lp")
            if stripped.startswith("Lp") and not stripped.isdigit():
                continue

            # W pozostałych wierszach (fragmentach nazwy, ilości, kodach, cenach).
            all_lines.append(stripped)

    # 2) Iteruj po all_lines sekwencyjnie i wyciągaj Lp, Name, Quantity, Barcode
    i = 0
    while i < len(all_lines):
        ln = all_lines[i]

        # 2a) Kod kreskowy
        if "Kod kres" in ln:
            parts = ln.split(":", maxsplit=1)
            if len(parts) == 2 and current:
                ean = parts[1].strip()
                # Przypisz kod EAN tylko jeśli bieżąca pozycja nie ma jeszcze swojego Barcode
                if not current.get("Barcode"):
                    current["Barcode"] = ean
            i += 1
            continue

        # 2b) Sama liczba: może to być Lp lub Quantity
        if re.fullmatch(r"\d+", ln):
            nxt = all_lines[i + 1] if (i + 1) < len(all_lines) else ""

            # 2b-i) Jeżeli następna linia to "szt." → to jest Quantity
            if nxt.lower() == "szt.":
                qty = int(ln)
                if current:
                    current["Quantity"] = qty
                    # Po przypisaniu ilości scal fragmenty nazwy w pełne pole Name
                    current["Name"] = " ".join(name_lines).strip()
                    name_lines = []
                    capture_name = False
                i += 2  # przeskocz również wiersz "szt."
                continue

            # 2b-ii) Jeżeli następna linia zawiera przynajmniej jedną literę i nie zaczyna się od "Kod kres" →
            #         to znaczy, że ln jest nowym Lp (numerem pozycji)
            elif re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", nxt) and not nxt.startswith("Kod kres"):
                Lp = int(ln)
                current = {"Lp": Lp, "Name": None, "Quantity": None, "Barcode": None}
                products.append(current)
                capture_name = True
                name_lines = []
                i += 1
                continue

            # 2b-iii) W przeciwnym razie traktujemy to jako „artefakt” (cena, numer w stopce lub inna liczba) → pomijamy
            else:
                i += 1
                continue

        # 2c) Fragment nazwy produktu: jeśli obecnie zbieramy nazwę (capture_name=True)
        if capture_name and ln:
            # Jeżeli linia wygląda jak cena ("123,45" lub "1 234,56"), pomiń:
            if re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", ln):
                i += 1
                continue
            # Jeżeli linia zaczyna się od "VAT", pomiń:
            if ln.startswith("VAT"):
                i += 1
                continue
            # Jeżeli linia to pojedynczy "/", pomiń:
            if ln == "/":
                i += 1
                continue
            # W innym razie dodajemy ln jako część nazwy produktu:
            name_lines.append(ln)
            i += 1
            continue

        # 2d) Pozostałe wiersze (np. puste) → pomiń
        i += 1

    # 3) Zbuduj DataFrame i odfilteruj wiersze bez Name lub bez Quantity
    df = pd.DataFrame(products)
    df = df.dropna(subset=["Name", "Quantity"]).reset_index(drop=True)
    return df


# ──────────────────────────────────────────────────────────────────────────────

# 1) File Uploader: wgraj PDF
uploaded_file = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if uploaded_file is None:
    st.info("Proszę wgrać plik PDF, aby uruchomić parser.")
    st.stop()

# 2) Wczytaj PDF za pomocą PyPDF2
try:
    pdf_reader = PyPDF2.PdfReader(uploaded_file)
except Exception as e:
    st.error(f"Nie udało się wczytać PDF-a: {e}")
    st.stop()

# 3) Parsowanie do DataFrame (z pokazaniem spinnera)
with st.spinner("Łączę strony i analizuję PDF…"):
    df = parse_pdf_to_dataframe(pdf_reader)

# 4) Wyświetl wynik w tabeli Streamlit
st.subheader("Wyekstrahowane pozycje zamówienia")
st.dataframe(df, use_container_width=True)

# 5) Przygotuj przycisk do pobrania jako Excel
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
