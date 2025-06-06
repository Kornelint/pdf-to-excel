# app.py

import streamlit as st
import pandas as pd
import re
import PyPDF2
import io

st.set_page_config(page_title="PDF → Excel", layout="wide")
st.title("Konwerter zamówienia PDF → Excel (obsługa dodatkowych układów)")

st.markdown(
    """
    Wgraj plik PDF ze zamówieniem. Aplikacja:
    1. Próbuje automatycznie wykryć, w jakim układzie jest PDF:
       - **Układ A** („Kod kres” w osobnej kolumnie, część wierszy rozbita między strony),
       - **Układ B** (EAN jest w tej samej linii co Lp, np. „1 5029040012366 Canagan Cat … 96,00 szt.”).
    2. Dla **Układu A** działa stary uniwersalny parser („Kod kres” na osobnej linii, nazwa rozbita przed/po cenach).
    3. Dla **Układu B** stosuje wzorzec pojedynczej linii:
       ```
       <Lp> <EAN> <pełna nazwa produktu> <ilość_liczba>,<liczba> szt. <…inne kolumny…>
       ```
       – w takim wierszu natychmiast wyciągamy:
         • Lp = pierwsza liczba,  
         • Barcode = druga liczba (zwykle 13-cyfrowa, EAN),  
         • Name = tekst między EAN a ilością (ciąg liter/nawiasów/spacji) aż do wiersza „<ilość>,<…> szt.”,  
         • Quantity = liczba przed „,00” i „szt.”.  
    4. W efekcie obsłużymy teraz 3 formaty:
       - PDF z “Kod kres” w osobnych wierszach (pierwsze dwa przykłady),
       - PDF z EAN w tej samej linii co Lp (np. “Wydruk.pdf”).
    5. Wynik: tabela z kolumnami `Lp`, `Name`, `Quantity`, `Barcode` i możliwość pobrania pliku Excel.
    """
)

def parse_layout_a(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla układu, gdzie Lp i nazwa mogą być w odrębnych wierszach,
    a "Kod kres.: <EAN>" występuje w osobnej linii.
    Kod EAN przypisujemy do ostatniego wcześniejszego Lp.
    Nazwę scalamy z fragmentów przed i po kolumnie cen.
    """
    # 1) Znajdź wszystkie indeksy Lp: linia czysta-liczba, a w poniższej linijce są litery (nie "szt." i nie cena i nie "Kod kres")
    idx_lp = []
    for i in range(len(all_lines) - 1):
        if re.fullmatch(r"\d+", all_lines[i]):
            nxt = all_lines[i + 1]
            if (
                re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", nxt)
                and nxt.lower() != "szt."
                and not re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", nxt)
                and not nxt.startswith("Kod kres")
            ):
                idx_lp.append(i)

    # 2) Znajdź wszystkie indeksy EAN (linię zaczynającą się od "Kod kres")
    idx_ean = [i for i, ln in enumerate(all_lines) if ln.startswith("Kod kres")]

    products = []
    for idx, lp_idx in enumerate(idx_lp):
        prev_lp = idx_lp[idx - 1] if idx > 0 else -1
        next_lp = idx_lp[idx + 1] if idx + 1 < len(idx_lp) else len(all_lines)

        # 2a) EAN: spośród wszystkich e in idx_ean, takich że prev_lp < e < next_lp, weź maksymalny e
        barcode = None
        valid_eans = [e for e in idx_ean if prev_lp < e < next_lp]
        if valid_eans:
            eidx = max(valid_eans)
            parts = all_lines[eidx].split(":", 1)
            if len(parts) == 2:
                barcode = parts[1].strip()

        # 2b) Nazwa + Quantity: od razu po lp_idx + 1 zbieramy fragmenty aż do wiersza <liczba> +"szt."
        name_parts = []
        qty = None
        qty_idx = None

        for j in range(lp_idx + 1, next_lp):
            ln = all_lines[j]
            # jeśli czysta liczba, a poniższy wiersz == "szt." → to ilość
            if re.fullmatch(r"\d+", ln) and (j + 1 < next_lp and all_lines[j + 1].lower() == "szt."):
                qty_idx = j
                qty = int(ln)
                break
            # w przeciwnym razie, jeśli ln zawiera litery (i nie jest cena, nie "VAT", nie "/") → fragment nazwy
            if (
                re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", ln)
                and not re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", ln)
                and not ln.startswith("VAT")
                and ln != "/"
                and not ln.startswith("ARA")
                and not ln.startswith("KAT")
            ):
                name_parts.append(ln)

        # jeśli nie znaleziono qty → pomiń
        if qty_idx is None:
            continue

        # 2c) Po znalezieniu qty, idź dalej aż do next_lp lub "Kod kres", dopisując fragmenty nazwy
        for k in range(qty_idx + 1, next_lp):
            ln2 = all_lines[k]
            if ln2.startswith("Kod kres"):
                break
            if (
                re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", ln2)
                and not re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", ln2)
                and not ln2.startswith("VAT")
                and ln2 != "/"
                and not ln2.startswith("ARA")
                and not ln2.startswith("KAT")
            ):
                name_parts.append(ln2)

        Name_val = " ".join(name_parts).strip()
        products.append({
            "Lp": int(all_lines[lp_idx]),
            "Name": Name_val,
            "Quantity": qty,
            "Barcode": barcode
        })

    df = pd.DataFrame(products)
    return df


def parse_layout_b(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla układu, w którym cała pozycja (Lp, EAN, nazwa, ilość) jest w jednej linii, np:
        1 5029040012366 Canagan Cat Can Chicken with Beef 75g 96,00 szt. 0,00 0,00
    W takim wierszu:
      - ^(\d+)\s+(\d{13})\s+(.+?)\s+(\d{1,3}),\d{2}\s+szt
        wyciągamy Lp, EAN, Name i Quantity.
      - Pozostałe dane (VAT, ceny netto/brutto) pomijamy.
    """
    products = []
    pattern = re.compile(
        r"^(\d+)\s+(\d{13})\s+(.+?)\s+(\d{1,3}),\d{2}\s+szt", 
        flags=re.IGNORECASE
    )
    for ln in all_lines:
        m = pattern.match(ln)
        if m:
            Lp_val = int(m.group(1))
            Barcode_val = m.group(2)
            Name_val = m.group(3).strip()
            Quantity_val = int(m.group(4).replace(" ", ""))
            products.append({
                "Lp": Lp_val,
                "Name": Name_val,
                "Quantity": Quantity_val,
                "Barcode": Barcode_val
            })

    df = pd.DataFrame(products)
    return df


def parse_pdf_generic(reader: PyPDF2.PdfReader) -> pd.DataFrame:
    """
    Miesza oba podejścia:
      1) Najpierw zbiera wszystkie wiersze (pomijając stopki i nagłówki) do listy all_lines.
      2) Następnie sprawdza, czy którykolwiek ln w all_lines pasuje do wzorca „Układ B” (Lp + EAN + nazwa w tej samej linii).
         - Jeśli tak przynajmniej raz, przyjmuje, że to Układ B i wywołuje parse_layout_b().
         - W przeciwnym razie wywołuje parse_layout_a().
    """
    all_lines = []
    started = False

    for page in reader.pages:
        raw = page.extract_text().split("\n")
        for ln in raw:
            stripped = ln.strip()

            # 1a) Stopki → jeśli znajdziemy "Wydrukowano", "Strona", lub "ZD <numer>", przerwij stronę
            if (
                stripped.startswith("Wydrukowano")
                or stripped.startswith("Strona")
                or re.match(r"ZD \d+/?\d*", stripped)
            ):
                break

            # 1b) Dopóki nie napotkamy nagłówka "Lp" lub "Nazwa towaru", pomijamy wiersze wstępu
            if not started:
                if (stripped.startswith("Lp") and not stripped.isdigit()) or stripped.lower().startswith("nazwa towaru"):
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

            # 1d) W przeciwnym razie dodajemy stripped do all_lines
            all_lines.append(stripped)

    # 2) Sprawdź, czy mamy przynajmniej jeden wiersz pasujący do Układu B (Lp + EAN w jednym ln)
    pattern_b = re.compile(r"^\d+\s+\d{13}\s+.+\s+\d{1,3},\d{2}\s+szt", flags=re.IGNORECASE)
    is_layout_b = any(pattern_b.match(ln) for ln in all_lines)

    if is_layout_b:
        return parse_layout_b(all_lines)
    else:
        return parse_layout_a(all_lines)


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

# 3) Parsowanie do DataFrame (pokazujemy spinner dla większych plików)
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
