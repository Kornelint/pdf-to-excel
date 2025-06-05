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
       - Stopki (linia zaczynająca się od "Strona", "Wydrukowano" lub zawierająca "ZD <numer>"),
       - Powtarzające się nagłówki tabeli (linie zaczynające się od "Lp" – niebędące czystą liczbą,
         linie zaczynające się od "Nazwa towaru", oraz wiersze: "Ilość", "J. miary", "Cena", 
         "Wartość", "netto", "brutto", "Indeks katalogowy").
    2. Na scalonym ciągu:
       - Znajdzie wszystkie miejsce, gdzie wiersz to czysta liczba a kolejny wiersz zawiera tekst (litery) – 
         to będzie ** Lp ** (numer pozycji).
       - Znajdzie wszystkie miejsce, gdzie wiersz zaczyna się od "Kod kres" (EAN).
       - Każdą linię "Kod kres.: <kod>" przypisze do tego **ostatniego wcześniej** Lp w scalonym ciągu.  
         Dzięki temu nawet jeśli kod EAN znajduje się na następnej stronie, wciąż trafi do odpowiedniego Lp.
    3. Dla każdej pozycji **Lp** (np. 14 czy 34) zbierze:
       - `Name`: wszystkie fragmenty tekstu od wiersza *zaraz po* Lp aż do momentu, gdy napotka wiersz 
         czysty-liczba + obok "szt." (czyli `Quantity`), a następnie — **aż do najbliższej linii "Kod kres".**  
         W rezultacie zbierzemy zarówno fragmenty nazwy zalegające przed kolumną cenową, jak i te – które w PDF-ie 
         pojawiają się dopiero za cenami (np. "indyk+tuńczyk").  
       - `Quantity`: wartość tej liczby (linia czysta-liczba, której kolejna linia to dokładnie `"szt."`).
       - `Barcode`: kod EAN pobrany z mapy utworzonej na podstawie `Kod kres.: <…>`.  
    4. Wyświetli tabelę w Streamlit i pozwoli pobrać wynikowy plik Excel z kolumnami:  
       `Lp`, `Name`, `Quantity`, `Barcode`.
    """
)

def parse_pdf_generic(reader: PyPDF2.PdfReader) -> pd.DataFrame:
    """
    Uniwersalny parser, obsługujący oba PDF-y (ten z nagłówkiem "Lp" i ten z nagłówkiem "Nazwa towaru").
    Całość działa w kilku krokach:

    ● 1) Zbieranie `all_lines`:
        - Przechodzimy po każdej stronie PDF-a, linia po linii.
        - Jeśli linia = stopka ("Wydrukowano"/"Strona"/"ZD <numer>"), przerywamy czytanie tej strony.
        - Dopóki nie napotkamy “pierwszego nagłówka” (linia zaczynająca się od "Lp" lub "Nazwa towaru"), ignorujemy linie wstępu.
        - Po wykryciu nagłówka:
            • Pomijamy powtarzające się nagłówki (kolejne linie zaczynające się od "Lp", "Nazwa towaru", 
              oraz „Ilość”, „J. miary”, „Cena”, „Wartość”, „netto”, „brutto”, „Indeks katalogowy”).
            • Każdą inną – dodajemy do listy `all_lines`, zachowując kolejność, w jakiej występują w dokumencie.

    ● 2) Znalezienie indeksów `idx_lp` (gdzie zaczyna się nowa pozycja Lp):
        - Szukamy wszystkich i takich i, że
            ```
            all_lines[i].isdigit() == True           # linia to czysta liczba
            and
            all_lines[i+1] zawiera dowolną literę     # wiersz obok to tekst (fragment nazwy)
            and nie (all_lines[i+1].lower() == "szt.") 
            and nie jest to linia-cena (np. "123,45").
            ```
        - Każde takie i to nowy słownik `current = {"Lp": int(all_lines[i]), "Name": None, "Quantity":None, "Barcode":None}` 
          i zapisujemy i w tablicy `idx_lp`.

    ● 3) Znalezienie indeksów `idx_ean` (wszystkie wiersze zaczynające się od "Kod kres"):
        - Dla każdej linii `j` w `all_lines`, jeśli `all_lines[j].startswith("Kod kres")`, wyciągamy ciąg po dwukropku:
            ```
            kod = all_lines[j].split(":",1)[1].strip()
            ```
          a następnie przypisujemy go do „ ostatniego wcześniej” Lp, czyli od szukamy
            ```
            candidates = [lp_i for lp_i in idx_lp if lp_i < j]
            lp_target = max(candidates)
            ean_map[lp_target] = kod
            ```
          W rezultacie nawet jeśli „Kod kres.: …” jest NA KOLEJNEJ stronie, a więc jego indeks j > indeksu Lp, 
          to i tak ten kod trafi do prawidłowej pozycji.

    ● 4) Budowanie pełnego rekordu produktu dla każdej pozycji `lp_idx ∈ idx_lp`:
        1. `Lp_val = int(all_lines[lp_idx])`
        2. Zbieranie fragmentów nazwy (`name_parts`):
            ```python
            name_parts = []
            j = lp_idx + 1
            qty_idx = None
            while j < n:
                ln = all_lines[j]
                #  ● jeśli znajdziemy ilość (czysta liczba, a poniższa linia to "szt."), to to jest `Quantity`
                if re.fullmatch(r"\d+", ln) and (j+1 < n and all_lines[j+1].lower() == "szt."):
                    qty_idx = j
                    break
                #  ● w przeciwnym razie, gdy ln zawiera litery (ORAZ nie wygląd jak cena, nie zaczyna się od "VAT", nie jest "/"), 
                #    to znaczy fragment nazwy, więc dorzuć go
                if re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", ln) \
                   and not re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", ln) \
                   and not ln.startswith("VAT") and ln != "/":
                    name_parts.append(ln)
                j += 1
            ```
            W efekcie:
            - Do momentu znalezienia ilości zbieramy fragmenty nazwy, które mogą występować ZARAZ po Lp (część przed cenami).
            - `qty_idx` to indeks w `all_lines`, gdzie jest czysta liczba, a poniżej `"szt."`.
        3. **Jeżeli nie znalazło `qty_idx` ⇒ pomijamy tę pozycję** (brak kompletnej pozycji).
        4. `Quantity_val = int(all_lines[qty_idx])`.
        5. **Zbieranie DALSZYCH fragmentów nazwy** do momentu, gdy napotkamy `"Kod kres"`, tzn.:
            ```python
            k = qty_idx + 1
            while k < n:
                ln2 = all_lines[k]
                if ln2.startswith("Kod kres"):
                    break
                #  ● Pomijajemy linie typu „ARA000XXX” czy „KATXXXXX” (to katalogowe ID), 
                #    pomijamy ceny ("123,45"/"1 234,56"), pomijamy "VAT", pomijamy pojedyncze "/".
                if re.fullmatch(r"[A-Z]{3}\d+", ln2) or re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", ln2) \
                   or ln2.startswith("VAT") or ln2 == "/":
                    k += 1
                    continue
                #  ● W przeciwnym razie, jeżeli w ln2 są litery, to dołączamy to do nazwy:
                if re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", ln2):
                    name_parts.append(ln2)
                k += 1
            ```
            Dzięki temu uchwycimy także te fragmenty nazwy, które w PDF-ie pojawiają się dopiero po kolumnach ilość/cena, np. “indyk+tuńczyk”.

        6. `Name_val = " ".join(name_parts).strip()`
        7. `Barcode_val = ean_map.get(lp_idx, None)` – może być `None`, jeżeli w PDF-ie w ogóle nie było linii „Kod kres” przypisanej do tej pozycji.

    ● 5) Na koniec tworzymy:
       ```python
       df = pd.DataFrame(products)
       df = df.dropna(subset=["Name","Quantity"]).reset_index(drop=True)
       ```
       i zwracamy `df` z kolumnami `Lp`, `Name`, `Quantity`, `Barcode`.

    Dzięki temu w jednym przebiegu ustrzelimy:
    - **Lp** (zarówno gdy PDF ma nagłówek „Lp”, jak i gdy ma nagłówek „Nazwa towaru”),
    - **pełną nazwę** (nawet gdy jest rozbita – fragment przed i fragment po cenach),
    - **ilość** (linia `<liczba>`, za którą stoi `"szt."`),
    - **EAN** (kodi kres., nawet jeśli trafia się na innej stronie).
    """

    # 1) Scal wszystkie strony w all_lines
    all_lines = []
    started = False
    for page in reader.pages:
        raw = page.extract_text().split("\n")
        for ln in raw:
            stripped = ln.strip()
            # a) jeżeli stopka (np. "Wydrukowano", "Strona" lub "ZD <numer>"), przerwij tę stronę:
            if stripped.startswith("Wydrukowano") or stripped.startswith("Strona") or re.match(r"ZD \d+/?\d*", stripped):
                break
            # b) dopóki nie zobaczymy nagłówka "Lp" lub "Nazwa towaru", pomijamy wiersze wstępu
            if not started:
                if (stripped.startswith("Lp") and not stripped.isdigit()) or stripped.lower().startswith("nazwa towaru"):
                    started = True
                    continue
                else:
                    continue
            # c) po wykryciu nagłówka, pomijamy powtórzone wiersze nagłówków:
            if stripped.startswith("Lp") and not stripped.isdigit():
                continue
            if stripped.lower().startswith("nazwa towaru"):
                continue
            if stripped in ["Ilość", "J. miary", "Cena", "Wartość", "netto", "brutto", "Indeks katalogowy", ""]:
                continue
            # d) w pozostałych wierszach (faktyczne dane tabeli) – dopisujemy:
            all_lines.append(stripped)

    n = len(all_lines)

    # 2) Znajdź wszystkie indeksy, w których jest "Lp" (czysta liczba + tekst obok)
    idx_lp = []
    for i in range(n - 1):
        if re.fullmatch(r"\d+", all_lines[i]):
            nxt = all_lines[i + 1]
            # kolejna linia zawiera litery (to fragment nazwy) i nie jest "szt." ani wierszem-ceną, ani "Kod kres"
            if (re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", nxt)
                    and nxt.lower() != "szt."
                    and not re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", nxt)
                    and not nxt.startswith("Kod kres")):
                idx_lp.append(i)

    # 3) Znajdź wszystkie indeksy z "Kod kres"
    idx_ean = [i for i, ln in enumerate(all_lines) if ln.startswith("Kod kres")]

    # 4) Zbuduj mapę {lp_idx: EAN}, przypisując każdemu kodowi EAN najbliższy wcześniejszy Lp
    ean_map = {}
    for e in idx_ean:
        parts = all_lines[e].split(":", 1)
        if len(parts) == 2:
            val = parts[1].strip()
            # wybieramy największy lp_idx < e
            candidates = [lp for lp in idx_lp if lp < e]
            if candidates:
                lp_t = max(candidates)
                ean_map[lp_t] = val

    # 5) Przejdź po każdej pozycji Lp i zbierz Name, Quantity, Barcode
    products = []
    for lp_idx in idx_lp:
        Lp_val = int(all_lines[lp_idx])

        # 5a) Zbierz fragmenty nazwy OD RAZU PO Lp aż do wiersza oznaczającego ilość
        name_parts = []
        j = lp_idx + 1
        qty_idx = None
        while j < n:
            ln = all_lines[j]
            # jeśli ta linia to czysta liczba, a kolejna to "szt." → mamy Quantity
            if re.fullmatch(r"\d+", ln) and (j + 1 < n and all_lines[j + 1].lower() == "szt."):
                qty_idx = j
                break
            # jeżeli ln zawiera litery i nie jest "cena" ani "VAT" ani "/" → kawałek nazwy
            if (re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", ln)
                    and not re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", ln)
                    and not ln.startswith("VAT")
                    and ln != "/"):
                name_parts.append(ln)
            j += 1

        # jeśli nie znaleziono qty_idx, pomijamy tę pozycję, bo nie kompletna
        if qty_idx is None:
            continue

        # 5b) Odczytaj Quantity jako int
        qty = int(all_lines[qty_idx])

        # 5c) Po odnalezieniu qty, dalej zbierzmy DODATKOWE fragmenty nazwy aż do "Kod kres"
        k = qty_idx + 1
        while k < n:
            ln2 = all_lines[k]
            # jeśli trafimy na "Kod kres", to wychodzimy
            if ln2.startswith("Kod kres"):
                break
            # pomijamy linie katalogowe np. "ARA000255" czy "KAT00208"
            if re.fullmatch(r"[A-Z]{3}\d+", ln2):
                k += 1
                continue
            # pomijamy ceny np. "123,45" czy "1 234,56", a także "VAT" lub "/"
            if re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", ln2) or ln2.startswith("VAT") or ln2 == "/":
                k += 1
                continue
            # jeśli w ln2 są litery → dołączamy do nazwy
            if re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", ln2):
                name_parts.append(ln2)
            k += 1

        # 5d) Scal nazwę i pobierz Barcode
        Name_val = " ".join(name_parts).strip()
        Barcode_val = ean_map.get(lp_idx, None)

        products.append({
            "Lp": Lp_val,
            "Name": Name_val,
            "Quantity": qty,
            "Barcode": Barcode_val
        })

    # 6) Zbuduj DataFrame i usuń wiersze, które nie mają ani Name, ani Quantity
    df = pd.DataFrame(products)
    df = df.dropna(subset=["Name", "Quantity"]).reset_index(drop=True)
    return df


# ──────────────────────────────────────────────────────────────────────────────

# 1) FileUploader – użytkownik wgrywa PDF
uploaded_file = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if uploaded_file is None:
    st.info("Proszę wgrać plik PDF, aby uruchomić parser.")
    st.stop()

# 2) Czytanie PDF przez PyPDF2
try:
    pdf_reader = PyPDF2.PdfReader(uploaded_file)
except Exception as e:
    st.error(f"Nie udało się wczytać PDF-a: {e}")
    st.stop()

# 3) Parsowanie do DataFrame (pokazujemy spinner, bo czasami plik jest duży)
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
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
