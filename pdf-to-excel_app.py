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
    1. Połączy wszystkie strony w jeden ciąg liniowy, pomijając stopki (linie zaczynające się od "Strona", "Wydrukowano", "ZD ...") 
       oraz powtórzone nagłówki tabeli.
    2. Wykryje, czy jest to **Format 1** (nagłówek „Lp.”) czy **Format 2** (nagłówek „Nazwa towaru lub usługa”).
    3. Dla **Formatu 1** wyłuska: `Lp → (fragmenty nazwy) → ilość (<liczba> + "szt.") → Kod kres.:<EAN>`.
    4. Dla **Formatu 2** wyłuska: `(jednoliniowa nazwa) → ilość (<liczba> + "szt.") → (inne kolumny, np. cena) → Kod kres.:<EAN>`.
    5. Wyświetli wynik w postaci tabeli oraz umożliwi pobranie pliku Excel 
       z kolumnami: `Lp`, `Name`, `Quantity`, `Barcode`.
    """
)

def parse_pdf_format1(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla formatu, w którym każda strona ma nagłówek "Lp.".
    all_lines: lista wszystkich linii tekstu po połączeniu stron i usunięciu stopki/nagłówków.
    Zwraca DataFrame z kolumnami ['Lp', 'Name', 'Quantity', 'Barcode'].
    """
    products = []
    current = None
    capture_name = False
    name_lines = []

    i = 0
    n = len(all_lines)
    while i < n:
        ln = all_lines[i]

        # 1) Jeśli linia zawiera "Kod kres", wyciągamy kod EAN z tekstu po ":" 
        if "Kod kres" in ln:
            parts = ln.split(":", maxsplit=1)
            if len(parts) == 2 and current:
                ean = parts[1].strip()
                # przypisz, jeśli current["Barcode"] nadal jest puste
                if not current.get("Barcode"):
                    current["Barcode"] = ean
            i += 1
            continue

        # 2) Jeżeli ln to sama liczba → może być to Lp lub Quantity
        if re.fullmatch(r"\d+", ln):
            nxt = all_lines[i + 1] if (i + 1) < n else ""

            # 2a) Jeżeli następna linia to "szt." → traktujemy ln jako Quantity
            if nxt.lower() == "szt.":
                qty = int(ln)
                if current:
                    current["Quantity"] = qty
                    current["Name"] = " ".join(name_lines).strip()
                    name_lines = []
                    capture_name = False
                i += 2  # przeskakujemy również linię "szt."
                continue

            # 2b) Jeżeli następna linia zawiera litery (rozpoczyna nazwę) i nie zaczyna się od "Kod kres" → to nowy Lp
            elif re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", nxt) and not nxt.startswith("Kod kres"):
                Lp_val = int(ln)
                current = {"Lp": Lp_val, "Name": None, "Quantity": None, "Barcode": None}
                products.append(current)
                capture_name = True
                name_lines = []
                i += 1
                continue

            # 2c) W przeciwnym razie uznajemy tę liczbę za artefakt (np. cena, fragment innej kolumny) i pomijamy
            else:
                i += 1
                continue

        # 3) Jeśli capture_name = True i ln zawiera litery → fragment nazwy produktu
        if capture_name and re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", ln):
            # pomijaj wiersze wyglądające jak cena, np. "123,45" lub "1 234,56"
            if re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", ln):
                i += 1
                continue
            # pomijaj linie zaczynające się od "VAT"
            if ln.startswith("VAT"):
                i += 1
                continue
            # pomijaj pojedynczy "/"
            if ln == "/":
                i += 1
                continue

            # W pozostałych przypadkach to fragment nazwy
            name_lines.append(ln)
            i += 1
            continue

        # 4) Wszelkie inne wiersze (puste, "Ilość", "J. miary" itp.) pomijamy
        i += 1

    # 5) Zbuduj DataFrame i odrzuć wiersze bez Name lub Quantity
    df = pd.DataFrame(products)
    df = df.dropna(subset=["Name", "Quantity"]).reset_index(drop=True)
    return df


def parse_pdf_format2(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla formatu z nagłówkiem "Nazwa towaru lub usługa" (np. "Gussto wola park.pdf").
    all_lines: lista wszystkich linii po usunięciu stopek i nagłówków.
    Zwraca DataFrame z kolumnami ['Lp', 'Name', 'Quantity', 'Barcode'], gdzie Lp pozostaje None.
    """
    products = []
    current = None

    i = 0
    n = len(all_lines)
    while i < n:
        ln = all_lines[i]

        # 1) Jeśli linia zawiera "Kod kres", wyciągnij EAN i przypisz do bieżącej pozycji
        if "Kod kres" in ln:
            parts = ln.split(":", maxsplit=1)
            if len(parts) == 2 and current and not current.get("Barcode"):
                current["Barcode"] = parts[1].strip()
            i += 1
            continue

        # 2) Jeśli ln to sama liczba i następna linia to "szt." → Quantity
        if re.fullmatch(r"\d+", ln):
            nxt = all_lines[i + 1] if (i + 1) < n else ""
            if nxt.lower() == "szt.":
                qty = int(ln)
                if current:
                    current["Quantity"] = qty
                i += 2
                continue
            else:
                i += 1
                continue

        # 3) Jeśli ln zawiera litery (nazwę produktu), a nie jest typowym nagłówkiem tabeli czy ceną:
        if re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", ln):
            # pomijaj linie "VAT", "Cena", "Wartość", "netto", "brutto" lub puste
            if ln.startswith("VAT") or ln.startswith("Cena") or ln.startswith("Wartość") or ln in ["netto", "brutto", ""]:
                i += 1
                continue
            # jeżeli ta linia wygląda jak cena (np. "123,45"), pomiń
            if re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", ln):
                i += 1
                continue

            # W przeciwnym razie: jest to nazwa nowego produktu
            current = {"Lp": None, "Name": ln.strip(), "Quantity": None, "Barcode": None}
            products.append(current)
            i += 1
            continue

        # 4) Inne wiersze pomijamy
        i += 1

    # 5) Odfiltruj wiersze bez Name lub Quantity
    df = pd.DataFrame(products)
    df = df.dropna(subset=["Name", "Quantity"]).reset_index(drop=True)
    return df


def parse_pdf_combined(reader: PyPDF2.PdfReader) -> pd.DataFrame:
    """
    Główna funkcja parsująca:
    1. Łączy wszystkie strony PDF w listę all_lines, pomijając:
       - każdą linię od pojawienia się "Wydrukowano", "Strona" lub "ZD <liczby>" (stopka),
       - dopóki nie napotka nagłówka "Lp." lub "Nazwa towaru lub usługa" – pomija wiersze wstępu,
       - powtórzone nagłówki tabel («Lp.» i «Nazwa towaru lub usługa», oraz wiersze nagłówków kolumn: 
         "Ilość", "J. miary", "Cena", "Wartość", "netto", "brutto").
    2. Rozpoznaje, czy wykryty nagłówek to format 1 czy format 2 (flaga detected_format1).
    3. Wywołuje odpowiedni parser (parse_pdf_format1 lub parse_pdf_format2).
    4. Zwraca pełny DataFrame.
    """
    all_lines = []
    started = False
    detected_format1 = False

    for page in reader.pages:
        raw = page.extract_text().split("\n")
        for ln in raw:
            stripped = ln.strip()

            # 1a) Pomijamy stopki: "Wydrukowano", "Strona", "ZD <liczby>"
            if stripped.startswith("Wydrukowano") or stripped.startswith("Strona") or re.match(r"ZD \d+/?\d*", stripped):
                break

            # 1b) Dopóki nie napotkamy nagłówka (Lp. lub Nazwa towaru), pomijamy
            if not started:
                if stripped.startswith("Lp.") and not stripped.isdigit():
                    started = True
                    detected_format1 = True
                    continue
                if stripped.lower().startswith("nazwa towaru"):
                    started = True
                    continue
                # dopóki nie ma headera, przeskakujemy
                continue

            # 1c) Po wykryciu headera, pomijamy wszystkie standardowe linie nagłówków:
            if stripped.startswith("Lp.") and not stripped.isdigit():
                continue
            if stripped.lower().startswith("nazwa towaru"):
                continue
            if stripped in ["Ilość", "J. miary", "Cena", "Wartość", "netto", "brutto", "Indeks katalogowy", ""]:
                continue

            # 1d) W pozostałych wierszach (dane tabeli) – dopisujemy do all_lines
            all_lines.append(stripped)

    # 2) Wybierz odpowiedni parser
    if detected_format1:
        return parse_pdf_format1(all_lines)
    else:
        return parse_pdf_format2(all_lines)


# ──────────────────────────────────────────────────────────────────────────────

# 1) FileUploader: wgraj PDF
uploaded_file = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if uploaded_file is None:
    st.info("Proszę wgrać plik PDF, aby uruchomić parser.")
    st.stop()

# 2) Spróbuj wczytać PDF za pomocą PyPDF2
try:
    pdf_reader = PyPDF2.PdfReader(uploaded_file)
except Exception as e:
    st.error(f"Nie udało się wczytać PDF-a: {e}")
    st.stop()

# 3) Parsowanie do DataFrame (ze spinnerem)
with st.spinner("Łączę strony i analizuję PDF…"):
    df = parse_pdf_combined(pdf_reader)

# 4) Wyświetlenie wynikowej tabeli
st.subheader("Wyekstrahowane pozycje zamówienia")
st.dataframe(df, use_container_width=True)

# 5) Przycisk do pobrania wyniku jako Excel
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
