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
       - Stopki (linie zaczynające się od "Strona", "Wydrukowano" lub zawierające "ZD <numer>"),
       - Powtarzające się nagłówki tabeli: każdą linię zaczynającą się od "Lp" (niebędącą czystą liczbą),
         każdą linię zaczynającą się od "Nazwa towaru", oraz wiersze: "Ilość", "J. miary", "Cena", "Wartość",
         "netto", "brutto", "Indeks katalogowy".
    2. Na scalonym ciągu (`all_lines`):
       - Zidentyfikuje wszystkie pozycje Lp – każdą linię, która jest czystą liczbą, a w kolejnej linii pojawiają się litery 
         (i ta kolejna linia nie jest "szt." ani wierszem ceny ani "Kod kres").  
       - Zlokalizuje wszystkie linie „Kod kres.: <EAN>” i utworzy listę indeksów `idx_ean`.
       - Podzieli `all_lines` na „interwały” oddzielone pozycjami Lp. Dla każdej pozycji Lp weźmie dokładnie te linie, 
         które między nią a następną pozycją Lp się znajdują. W tej grupie:
         • znajdzie tzw. `qty_idx` (pierwszą linię będącą czystą liczbą, której kolejna linia to "szt.") → to jest `Quantity`,  
         • zbierze fragmenty nazwy – wszystkie wiersze zawierające litery (nie wyglądające jak cena, nie zaczynające się od “VAT”, nie będące “/”, 
           nie zaczynające się od “ARA” ani “KAT”), zarówno _przed_ jak i _po_ kolumnie ilości/ceny, aż do momentu napotkania linii “Kod kres”.  
         • spośród `idx_ean` wybierze ten indeks `e` (jeśli istnieje), który leży w przedziale `(poprzednie_Lp, następne_Lp)` i jest największy (czyli EAN dla tej pozycji).  
       • Sklei pełną nazwę, ustawi `Quantity`, pobierze `Barcode`.  
    3. Wyświetli wynik jako tabelę z kolumnami:
       `Lp`, `Name`, `Quantity`, `Barcode`  
       oraz umożliwi pobranie pliku Excel, zawierającego te cztery kolumny.
    """
)


def parse_pdf_generic(reader: PyPDF2.PdfReader) -> pd.DataFrame:
    """
    Uniwersalny parser, który poradzi sobie z PDF-ami zawierającymi:
      - format z nagłówkiem "Lp." 
      - lub format z nagłówkiem "Nazwa towaru lub usługa" (choć ostatecznie oba wczytujemy tą samą logiką).
    Kluczowe etapy:
      1) Scalanie stron w `all_lines`, pomijając stopki i powtarzające się nagłówki.
      2) Wyszukiwanie indeksów Lp (czysta liczba + poniżej linia z literami).
      3) Wyszukiwanie indeksów EAN (linia zaczynająca się od "Kod kres").
      4) Dla każdego Lp pobieranie:
         - `Name`: wszystkie wiersze z literami (przed i po cenach) aż do napotkania linii "Kod kres",
         - `Quantity`: ta liczba, której kolejna linia to "szt.",
         - `Barcode`: EAN wybrany spośród wszystkich “Kod kres” leżących między tą pozycją Lp a następnym Lp.
      5) Zwraca pandas.DataFrame z kolumnami ['Lp', 'Name', 'Quantity', 'Barcode'].
    """

    # 1) Scal wszystkie strony w all_lines, pomijając stopy i nagłówki
    all_lines = []
    started = False

    for page in reader.pages:
        raw = page.extract_text().split("\n")
        for ln in raw:
            stripped = ln.strip()

            # 1a) Jeśli to stopka → przerywamy tę stronę
            if (
                stripped.startswith("Wydrukowano")
                or stripped.startswith("Strona")
                or re.match(r"ZD \d+/?\d*", stripped)
            ):
                break

            # 1b) Dopóki nie napotkamy nagłówka "Lp" lub "Nazwa towaru", pomijamy
            if not started:
                if (
                    (stripped.startswith("Lp") and not stripped.isdigit())
                    or stripped.lower().startswith("nazwa towaru")
                ):
                    started = True
                    continue
                else:
                    continue

            # 1c) Po wykryciu nagłówka, pomijamy powtórzone wiersze nagłówków tabeli:
            if (stripped.startswith("Lp") and not stripped.isdigit()) or stripped.lower().startswith("nazwa towaru"):
                continue
            if (
                stripped.lower().startswith("ilo")
                or stripped.lower().startswith("j. miary")
                or stripped.lower().startswith("cena")
                or stripped.lower().startswith("warto")
                or stripped.lower().startswith("netto")
                or stripped.lower().startswith("brutto")
                or stripped.lower().startswith("indeks katalogowy")
                or stripped == ""
            ):
                continue

            # 1d) W przeciwnym razie to faktyczne dane → dodajemy do all_lines
            all_lines.append(stripped)

    n = len(all_lines)

    # 2) Znajdź indeksy Lp: linia = czysta liczba, a poniżej wiersz zawiera litery (i nie jest "szt." ani cena, ani "Kod kres")
    idx_lp = []
    for i in range(n - 1):
        line = all_lines[i]
        nxt = all_lines[i + 1]
        if re.fullmatch(r"\d+", line):
            # poniżej musi być tekst (fragment nazwy), a nie "szt." i nie wiersz-cena, nie "Kod kres"
            if (
                re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", nxt)
                and nxt.lower() != "szt."
                and not re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", nxt)
                and not nxt.startswith("Kod kres")
            ):
                idx_lp.append(i)

    # 3) Znajdź indeksy wszystkich "Kod kres"
    idx_ean = [i for i, ln in enumerate(all_lines) if ln.startswith("Kod kres")]

    # 4) Dla każdej pozycji Lp zbierz Name, Quantity i Barcode
    products = []
    for idx, lp_idx in enumerate(idx_lp):
        prev_lp = idx_lp[idx - 1] if idx > 0 else -1
        next_lp = idx_lp[idx + 1] if idx + 1 < len(idx_lp) else n

        # 4a) Wybierz EAN w przedziale między poprzednim lp a kolejnym lp
        ean = None
        valid_eans = [e for e in idx_ean if prev_lp < e < next_lp]
        if valid_eans:
            # bierzemy ten z największym indeksem (czyli najbliższy Lp od dołu)
            e = max(valid_eans)
            parts = all_lines[e].split(":", 1)
            if len(parts) == 2:
                ean = parts[1].strip()

        # 4b) Zbierz fragmenty nazwy i ilość
        name_parts = []
        qty = None
        qty_idx = None

        # ● Najpierw kolejne linie aż do wiersza "ilość":
        for j in range(lp_idx + 1, next_lp):
            ln = all_lines[j]
            # jeśli trafimy na "ilość" (liczba, a poniższa linia to dokładnie "szt.") → to Quantity
            if re.fullmatch(r"\d+", ln) and (j + 1 < next_lp and all_lines[j + 1].lower() == "szt."):
                qty_idx = j
                qty = int(ln)
                break

            # w przeciwnym razie, jeżeli to linia z literami, ale nie wygląda jak cena, nie zaczyna się od "VAT", nie jest "/" i nie jest kodem katalogowym "ARA..." lub "KAT..."
            if (
                re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", ln)
                and ln.lower() != "szt."
                and not re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", ln)
                and not ln.startswith("VAT")
                and ln != "/"
                and not ln.startswith("ARA")
                and not ln.startswith("KAT")
            ):
                name_parts.append(ln)

        # Jeśli nie znaleziono ilości → pomiń
        if qty_idx is None:
            continue

        # 4c) Po odczytaniu ilości, dalej zbieramy kolejny fragment nazwy aż do momentu, gdy napotkamy "Kod kres"
        for k in range(qty_idx + 1, next_lp):
            ln2 = all_lines[k]
            if ln2.startswith("Kod kres"):
                # natrafiliśmy na linię z EAN; przerwijmy zbieranie nazwy
                break

            # jeżeli to nazwa (litery) i nie wygląda jak cena, nie zaczyna się od "VAT", nie jest "/" i nie jest "ARA..." ani "KAT..."
            if (
                re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", ln2)
                and ln2.lower() != "szt."
                and not re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", ln2)
                and not ln2.startswith("VAT")
                and ln2 != "/"
                and not ln2.startswith("ARA")
                and not ln2.startswith("KAT")
            ):
                name_parts.append(ln2)

        # 4d) Scalona nazwa i zapis do listy produktów
        Name_val = " ".join(name_parts).strip()
        products.append(
            {
                "Lp": int(all_lines[lp_idx]),
                "Name": Name_val,
                "Quantity": qty,
                "Barcode": ean,
            }
        )

    # 5) Zbuduj DataFrame i usuń wiersze brakujące nazwy lub ilości
    df = pd.DataFrame(products)
    df = df.dropna(subset=["Name", "Quantity"]).reset_index(drop=True)
    return df


# ──────────────────────────────────────────────────────────────────────────────

# 1) FileUploader – użytkownik wgrywa PDF
uploaded_file = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if uploaded_file is None:
    st.info("Proszę wgrać plik PDF, aby uruchomić parser.")
    st.stop()

# 2) Wczytanie PDF przez PyPDF2
try:
    pdf_reader = PyPDF2.PdfReader(uploaded_file)
except Exception as e:
    st.error(f"Nie udało się wczytać PDF-a: {e}")
    st.stop()

# 3) Parsowanie do DataFrame (pokazywany jest spinner podczas przetwarzania)
with st.spinner("Łączę strony i analizuję PDF…"):
    df = parse_pdf_generic(pdf_reader)

# 4) Wyświetlenie wynikowej tabeli w Streamlit
st.subheader("Wyekstrahowane pozycje zamówienia")
st.dataframe(df, use_container_width=True)

# 5) Przycisk do pobrania pliku Excel
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
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
