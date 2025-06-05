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
    1. Przetworzy PDF, scalając wszystkie strony w jeden ciąg liniowy,  
       pomijając stopki (linie zaczynające się od "Strona", "Wydrukowano", "ZD ...") i powtórzone nagłówki tabeli.  
    2. Zidentyfikuje, który format PDF mamy do czynienia:
       - **Format 1** (np. „Zamówienie nr ZD 0175_05_25.pdf”): posiada kolumnę „Lp” w nagłówku i później bloki w układzie  
         `Lp → nazwa (w kilku wierszach) → ilość („<liczba>” + „szt.”) → kod kres.: <EAN>`.  
       - **Format 2** (np. „Gussto wola park.pdf”): posiada nagłówek „Nazwa towaru lub usługa” i później bloki w układzie  
         `nazwa (jedna linia) → ilość („<liczba>” + „szt.”) → (inne kolumny, często cena) → kod kres.: <EAN>`.  
    3. Na podstawie wykrytego formatu zastosuje odpowiedni parser, który:
       - zbierze wszystkie produkty (łącznie z tymi rozbitymi między stronami),  
       - zeskle sensownie fragmenty nazw,  
       - wyciągnie ilość i kod EAN (nawet gdy „Kod kres.” znalazł się na następnej stronie).  
    4. Wyświetli wynik w formie tabeli oraz umożliwi pobranie pliku Excel z kolumnami:  
       `Lp (jeśli jest), Name, Quantity, Barcode`.
    """
)

def parse_pdf_format1(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla formatu z kolumną 'Lp'. 
    all_lines: lista linii po scaleniu wszystkich stron (bez stopek/nagłówków).
    Struktura bloków:
      Lp                 → linia czysto cyfrowa, a kolejna linia to albo fragment nazwy, albo inna cyfra (np. gdy nazwa zaczyna się cyfrą – rzadziej).
      Nazwa              → od momentu po Lp zbieramy kolejne nie-puste linie aż do napotkania czystej liczby, która jest ilością.
      Ilość + 'szt.'    → linia z liczbą, a następna to dokładnie "szt." → to jest Quantity.
      Kod kres.: <EAN>  → linia z frazą "Kod kres." → wyciągamy EAN.
    Logika:
      1. Przechodzimy kolejno po all_lines.  
      2. Jeśli linia to sama liczba i następna linia to nie „szt.” → traktujemy to jako nowy Lp.  
      3. Gdy capture_name=True i linia zawiera litery → to fragment nazwy.  
      4. Gdy linia to sama liczba a kolejna linia to „szt.” → to Quantity, scal nazwy w `Name`.  
      5. Gdy linia zawiera "Kod kres" → wyciągnij wszystko po dwukropku jako `Barcode`.  
    """
    products = []
    current = None
    capture_name = False
    name_lines = []

    i = 0
    n = len(all_lines)
    while i < n:
        ln = all_lines[i]

        # 1) Jeśli linia zawiera "Kod kres", wyciągamy EAN:
        if "Kod kres" in ln:
            parts = ln.split(":", maxsplit=1)
            if len(parts) == 2 and current:
                ean = parts[1].strip()
                if not current.get("Barcode"):
                    current["Barcode"] = ean
            i += 1
            continue

        # 2) Jeżeli ln to sama liczba:
        if re.fullmatch(r"\d+", ln):
            nxt = all_lines[i + 1] if (i + 1) < n else ""

            # 2a) Jeśli następna linia to "szt." → ln to Quantity
            if nxt.lower() == "szt.":
                qty = int(ln)
                if current:
                    current["Quantity"] = qty
                    current["Name"] = " ".join(name_lines).strip()
                    name_lines = []
                    capture_name = False
                i += 2
                continue

            # 2b) Jeżeli następna linia zawiera litery i nie zaczyna się od "Kod kres" → to nowy Lp
            elif re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", nxt) and not nxt.startswith("Kod kres"):
                Lp_val = int(ln)
                current = {"Lp": Lp_val, "Name": None, "Quantity": None, "Barcode": None}
                products.append(current)
                capture_name = True
                name_lines = []
                i += 1
                continue

            # 2c) W przeciwnym razie – pomijamy (prawdopodobnie liczba to fragment tabeli np. cena)
            else:
                i += 1
                continue

        # 3) Jeżeli capture_name=True i ln zawiera choć jedną literę → to fragment nazwy
        if capture_name and re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", ln):
            # Ignorujemy wiersze wyglądające jak cena: "123,45" lub "1 234,56"
            if re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", ln):
                i += 1
                continue
            # Ignorujemy linię zaczynającą się od "VAT"
            if ln.startswith("VAT"):
                i += 1
                continue
            # Ignorujemy pojedynczy "/"
            if ln == "/":
                i += 1
                continue

            # W pozostałych przypadkach to fragment nazwy
            name_lines.append(ln)
            i += 1
            continue

        # 4) Inne wiersze (np. puste, ceny, "Ilosc", "J. miary" itp.) – pomiń
        i += 1

    # Po zakończeniu pętli: zbuduj DataFrame i odrzuć wiersze bez Name/Quantity
    df = pd.DataFrame(products)
    df = df.dropna(subset=["Name", "Quantity"]).reset_index(drop=True)
    return df


def parse_pdf_format2(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla formatu z nagłówkiem "Nazwa towaru lub usługa" (np. „Gussto wola park.pdf”).
    Struktura bloków:
      Nazwa           → jednolinijkowa nazwa produktu.
      Ilość + "szt." → linia z liczbą i bezpośrednio "szt." w następnej linii.
      Następne kolumny (cena, VAT itp.) pomijamy.
      Kod kres.:<EAN> → linia, w której pojawia się "Kod kres:" → wyciągamy EAN.
    Logika:
      1. Przechodzimy sekwencyjnie po all_lines.
      2. Jeśli ln to czysty tekst zawierający litery i nie przypomina ceny ani "VAT", a capture_name=False:
         traktujemy go jako nazwa nowego produktu, tworzymy nowy słownik current (bez Lp).
      3. Jeżeli następną linia to "szt.", a ln jest liczbą → tratimy to jako Quantity.
      4. Jeżeli ln zawiera "Kod kres", wyciągamy EAN i przypisujemy do current.
      5. Powtarzamy aż do końca all_lines.
    """
    products = []
    current = None

    i = 0
    n = len(all_lines)
    while i < n:
        ln = all_lines[i]

        # 1) Jeśli ln zawiera "Kod kres" → to jest EAN dla current
        if "Kod kres" in ln:
            parts = ln.split(":", maxsplit=1)
            if len(parts) == 2 and current and not current.get("Barcode"):
                current["Barcode"] = parts[1].strip()
            i += 1
            continue

        # 2) Jeżeli ln to sama liczba i następna linia to "szt." → to Quantity
        if re.fullmatch(r"\d+", ln):
            nxt = all_lines[i + 1] if (i + 1) < n else ""
            if nxt.lower() == "szt.":
                qty = int(ln)
                if current:
                    current["Quantity"] = qty
                i += 2
                continue
            # W przeciwnym razie liczba bez "szt." to może być cena lub fragment stopki, pomiń
            i += 1
            continue

        # 3) Jeżeli ln zawiera litery i nie jest typowym nagłówkiem (np. "VAT", "Cena", pusta linia) →
        #    traktujemy go jako nazwa nowego produktu
        if re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", ln):
            # Ignorujemy linie "VAT", "Cena", "Wartość", puste, "netto", "brutto"
            if ln.startswith("VAT") or ln.startswith("Cena") or ln.startswith("Wartość") \
               or ln in ["netto", "brutto", ""] or re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", ln):
                i += 1
                continue

            # To powinno być początek nowego produktu:
            current = {"Lp": None, "Name": ln.strip(), "Quantity": None, "Barcode": None}
            products.append(current)
            i += 1
            continue

        # 4) Wszelkie inne wiersze – pomiń
        i += 1

    # Teraz usuń wiersze bez Quantity (bo każdy produkt musi mieć ilość) lub bez Name
    df = pd.DataFrame(products)
    df = df.dropna(subset=["Name", "Quantity"]).reset_index(drop=True)
    return df


def parse_pdf_combined(reader: PyPDF2.PdfReader) -> pd.DataFrame:
    """
    Funkcja nadrzędna, która:
    1. Scala wszystkie strony w all_lines, pomijając:
       - każdą linię od pojawienia się "Wydrukowano" / "Strona" / "ZD <liczby>" (stopka),
       - dopóki nie natrafimy na nagłówek „Nazwa towaru” lub „Lp”, wstrzymuje zbieranie (pomija wiersze wstępu),
       - pomija powtórzone nagłówki tabeli („Lp ...” i „Nazwa towaru lub usługa”, „Ilość”, „J. miary”, „Cena”, „Wartość”, „netto”, „brutto”).
    2. Sprawdza, czy w all_lines występuje jakakolwiek linia zaczynająca się od „Lp” (dokładnie) –
       jeśli tak → wybiera parse_pdf_format1; w przeciwnym razie → parse_pdf_format2.
    3. Zwraca DataFrame.
    """
    raw_lines = []
    started = False

    # 1) Zbierz wszystkie strony w all_lines
    for page in reader.pages:
        page_lines = page.extract_text().split("\n")
        for ln in page_lines:
            stripped = ln.strip()

            # 1a) Pomijamy stopki: „Wydrukowano”, „Strona”, linie zaczynające się od "ZD " (numer zamówienia w stopce)
            if stripped.startswith("Wydrukowano") or stripped.startswith("Strona") or re.match(r"ZD \d+/", stripped):
                break

            # 1b) Dopóki nie napotkamy jednego z nagłówków („Lp” lub „Nazwa towaru”), wstrzymujemy zbieranie
            if not started:
                # Gdy znajdziemy linię zaczynającą się od "Lp" nie będącą czystą liczbą:
                if stripped.startswith("Lp") and not stripped.isdigit():
                    started = True
                    # Pomijamy tę linię nagłówka tabeli („Lp    Nazwa    Ilość”)
                    continue
                # Gdy znajdziemy linię zaczynającą się od "Nazwa towaru" (drugi format):
                if stripped.lower().startswith("nazwa towaru"):
                    started = True
                    # Pomijamy ją, bo to nagłówek
                    continue
                # W przeciwnym razie przeskoczemy (zwykle nagłówek PDF, logo, adres, itp.)
                continue

            # 1c) Jeśli już started=True, pomijamy typowe „drugi rząd nagłówka tabeli” 
            #     („Ilość”, „J. miary”, „Cena”, „Wartość”, „netto”, „brutto”)
            if stripped in ["Ilość", "J. miary", "Cena", "Wartość", "netto", "brutto", ""]:
                continue

            # 1d) W przeciwnym wypadku dopisujemy do raw_lines
            raw_lines.append(stripped)

    # 2) Rozpoznaj format: jeśli w raw_lines pojawi się co najmniej jedno "Lp    " (tj. linia zaczynająca się od "Lp" i nie będąca liczbą):
    has_lp_header = any(ln.startswith("Lp") for ln in raw_lines)

    if has_lp_header:
        # Format 1: parser z Lp
        df = parse_pdf_format1(raw_lines)
    else:
        # Format 2: parser z nagłówkiem nazwy towaru
        df = parse_pdf_format2(raw_lines)

    return df


# ──────────────────────────────────────────────────────────────────────────────

# 1) FileUploader: wczytaj PDF
uploaded_file = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if uploaded_file is None:
    st.info("Proszę wgrać plik PDF, aby uruchomić parser.")
    st.stop()

# 2) Spróbuj wczytać PDF przez PyPDF2
try:
    pdf_reader = PyPDF2.PdfReader(uploaded_file)
except Exception as e:
    st.error(f"Nie udało się wczytać PDF-a: {e}")
    st.stop()

# 3) Parsowanie do DataFrame
with st.spinner("Łączę strony i analizuję PDF…"):
    df = parse_pdf_combined(pdf_reader)

# 4) Wyświetlenie wyniku
st.subheader("Wyekstrahowane pozycje zamówienia")
st.dataframe(df, use_container_width=True)

# 5) Pobranie wyniku jako Excel
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
