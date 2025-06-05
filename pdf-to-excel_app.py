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
       - Stopki, czyli linie zaczynające się od "Strona", "Wydrukowano" lub zawierające "ZD <liczby>",
       - Powtórzone nagłówki tabeli: każdą linię zaczynającą się od "Lp" (nie będącą czystą liczbą),
         każdą linię zaczynającą się od "Nazwa towaru", oraz wiersze: "Ilość", "J. miary", "Cena", "Wartość", "netto", "brutto", "Indeks katalogowy".
    2. Ze scalonego ciągu `all_lines`:
       - Zidentyfikuje każdy wiersz, który jest czystą liczbą, a kolejny wiersz zawiera litery (i nie jest „szt.” ani nie wygląda jak cena) — to będzie **Lp** (numer pozycji).
       - Znajdzie wszystkie indeksy wierszy zawierające „Kod kres” — to są pozycje z kodem EAN.
       - Każdą linię „Kod kres.: <EAN>” przypisze do najbliższej **wcześniejszej** (poprzedzającej) pozycji Lp, jeżeli taka istnieje.
       - Dla każdej pozycji Lp utworzy:
         • `Lp` (int),  
         • `Name` — scalone w jeden ciąg wszystkie kolejne wiersze zawierające litery (pomijając wiersze wyglądające jak cena, „VAT” czy „/”), aż do napotkania ilości,  
         • `Quantity` (int) z wiersza, gdzie wiersz to czysta liczba, a kolejny wiersz to dokładnie `"szt."`,  
         • `Barcode` (EAN) wyciągnięty z mapy kodów EAN utworzonej wcześniej.
    3. Wyświetli wynik w interaktywnej tabeli oraz umożliwi pobranie pliku Excel z kolumnami:  
       `Lp`, `Name`, `Quantity`, `Barcode`.
    """
)


def parse_pdf_generic(reader: PyPDF2.PdfReader) -> pd.DataFrame:
    """
    Uniwersalny parser, który wczytuje wszystkie strony PDF w listę `all_lines`,
    usuwa nagłówki i stopki, a następnie:

    1. Wyszukuje indeksy wierszy będące numerami pozycji (Lp): to czysta liczba, po której następuje wiersz z literami
       (nie będący "szt." ani wierszem ceny).
    2. Wyszukuje indeksy wierszy zawierające frazę "Kod kres" i tworzy mapę {indeks_Lp: kod_EAN}
       przypisując każdemu kodowi EAN najbliższy wcześniejszy Lp.
    3. Dla każdej pozycji Lp zbiera kolejne wiersze z literami jako fragmenty nazwy, aż do linii, która jest ilością
       (czysta liczba i następna linia to "szt."), i zapisuje `Quantity`.
    4. Łączy wszystko w listę słowników: {"Lp":…, "Name":…, "Quantity":…, "Barcode":…} i zwraca jako DataFrame.
    """
    all_lines = []
    started = False

    # 1) Scal wszystkie strony, pomijając stopki i początkowe nagłówki
    for page in reader.pages:
        raw = page.extract_text().split("\n")
        for ln in raw:
            stripped = ln.strip()

            # 1a) Stopki: "Wydrukowano", "Strona", lub linie typu "ZD 0175/05"
            if stripped.startswith("Wydrukowano") or stripped.startswith("Strona") or re.match(r"ZD \d+/?\d*", stripped):
                break

            # 1b) Dopóki nie napotkamy nagłówka tabeli, pomijamy wiersze wstępu:
            #     nagłówek "Lp" lub "Nazwa towaru"
            if not started:
                if (stripped.startswith("Lp") and not stripped.isdigit()) or stripped.lower().startswith("nazwa towaru"):
                    started = True
                    continue
                # brak nagłówka → pomiń
                continue

            # 1c) Po wykryciu nagłówka pomijamy powtórzone linie nagłówków:
            if stripped.startswith("Lp") and not stripped.isdigit():
                continue
            if stripped.lower().startswith("nazwa towaru"):
                continue
            if stripped in ["Ilość", "J. miary", "Cena", "Wartość", "netto", "brutto", "Indeks katalogowy", ""]:
                continue

            # 1d) W każdym innym wierszu (rzeczywiste dane tabeli) dodajemy stripped do all_lines
            all_lines.append(stripped)

    # 2) Zidentyfikuj indeksy Lp: czysta liczba, a następny wiersz zawiera litery (i nie jest "szt." ani cena)
    idx_lp = []
    n = len(all_lines)
    for i in range(n - 1):
        line = all_lines[i]
        nxt = all_lines[i + 1]
        if re.fullmatch(r"\d+", line):
            # następny wiersz zawiera litery (to początek nazwy) i nie jest "szt." ani wierszem-ceny
            if re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", nxt) and nxt.lower() != "szt." \
               and not re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", nxt) and not nxt.startswith("Kod kres"):
                idx_lp.append(i)

    # 3) Zlokalizuj indeksy wierszy zawierające "Kod kres"
    idx_ean = [i for i, ln in enumerate(all_lines) if ln.startswith("Kod kres")]

    # 4) Dla każdego kodu EAN przypisz go do najbliższego wcześniejszego Lp
    ean_map = {}  # {lp_index: kod_EAN}
    for e in idx_ean:
        parts = all_lines[e].split(":", 1)
        if len(parts) == 2:
            val = parts[1].strip()
            # znajdź największy lp_idx < e
            candidates = [lp for lp in idx_lp if lp < e]
            if candidates:
                lp_target = max(candidates)
                ean_map[lp_target] = val

    # 5) Dla każdej pozycji Lp zbuduj pełen rekord
    products = []
    for lp_idx in idx_lp:
        Lp_val = int(all_lines[lp_idx])

        # 5a) Zbierz fragmenty nazwy od lp_idx+1 do wiersza oznaczającego ilość
        name_parts = []
        j = lp_idx + 1
        qty_idx = None
        while j < n:
            ln = all_lines[j]
            # jeśli linia to czysta liczba i j+1 < n oraz all_lines[j+1].lower() == "szt." → to Quantity
            if re.fullmatch(r"\d+", ln) and (j + 1 < n and all_lines[j + 1].lower() == "szt."):
                qty_idx = j
                break

            # jeżeli linia zawiera litery, nie wygląda jak cena, nie zaczyna się od "VAT" i nie jest "/" → to fragment nazwy
            if re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", ln) \
               and not re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", ln) \
               and not ln.startswith("VAT") and ln != "/":
                name_parts.append(ln)

            j += 1

        # jeśli nie znaleziono qty_idx, pomiń tę pozycję
        if qty_idx is None:
            continue

        # 5b) Odczytaj Quantity
        qty = int(all_lines[qty_idx])

        # 5c) Scal wszystkie fragmenty nazwy w jeden ciąg
        Name_val = " ".join(name_parts).strip()

        # 5d) Pobierz Barcode z ean_map (jeśli istnieje dla tego lp_idx)
        Barcode_val = ean_map.get(lp_idx)

        products.append({
            "Lp": Lp_val,
            "Name": Name_val,
            "Quantity": qty,
            "Barcode": Barcode_val
        })

    # 6) Zbuduj DataFrame i odrzuć wiersze bez Name lub bez Quantity
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

# 4) Wyświetlenie rezultatu w tabeli
st.subheader("Wyekstrahowane pozycje zamówienia")
st.dataframe(df, use_container_width=True)

# 5) Przycisk do pobrania wyniku jako plik Excel (.xlsx)
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
