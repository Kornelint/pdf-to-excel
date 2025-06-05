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
       - Stopki, czyli linie zaczynające się od "Strona", "Wydrukowano" lub "ZD <liczby>".
       - Powtórzone nagłówki tabeli: każdą linię zaczynającą się od "Lp" (nie będącą czystą liczbą),
         linię zaczynającą się od "Nazwa towaru" oraz wiersze: "Ilość", "J. miary", "Cena", "Wartość", "netto", "brutto", "Indeks katalogowy".
    2. Po scaleniu wszystkich stron w listę `all_lines`:
       - Znajdzie wszystkie indeksy linii, które są czystą liczbą, a kolejna linia zawiera przynajmniej jedną literę – to jest **Lp** (numer pozycji).
       - Znajdzie wszystkie indeksy linii, które zawierają frazę "Kod kres" – to są pozycje z kodem EAN.
         Kod EAN przypisuje do najbliższego **następnego** numeru pozycji (Lp), jeżeli taki istnieje.
       - Dla każdej pozycji Lp zbuduje:
         • `Lp` (int),  
         • `Name` – łańcuch scalony ze wszystkich kolejnych wierszy zawierających litery (z pominięciem cen, VAT, "/") aż do napotkania ilości,  
         • `Quantity` (int) z wiersza, gdzie linia to czysta liczba, a następna linia to "szt.",  
         • `Barcode` (EAN) wyciągnięty ze skojarzonej linii "Kod kres." (jeśli taka istniała).
    3. Wyświetli wynik w postaci interaktywnej tabeli oraz umożliwi pobranie pliku `.xlsx` z kolumnami:
       `Lp`, `Name`, `Quantity`, `Barcode`.
    """
)


def parse_pdf_generic(reader: PyPDF2.PdfReader) -> pd.DataFrame:
    """
    Uniwersalny parser wyodrębniający wszystkie pozycje z PDF-a (formaty z Lp i różne layouty).
    Zwraca DataFrame z kolumnami ['Lp', 'Name', 'Quantity', 'Barcode'].
    """
    # Krok 1: scalamy wszystkie strony w all_lines, usuwając stopki i powtarzające się nagłówki
    all_lines = []
    started = False

    for page in reader.pages:
        raw = page.extract_text().split("\n")
        for ln in raw:
            stripped = ln.strip()

            # 1a) Jeśli to stopka: linia zaczyna się od "Wydrukowano", "Strona" lub "ZD <liczby>"
            if stripped.startswith("Wydrukowano") or stripped.startswith("Strona") or re.match(r"ZD \d+/?\d*", stripped):
                break

            # 1b) Dopóki nie napotkamy nagłówka ("Lp" lub "Nazwa towaru"), ignorujemy wiersze wstępu
            if not started:
                if (stripped.startswith("Lp") and not stripped.isdigit()) or stripped.lower().startswith("nazwa towaru"):
                    # znaleźliśmy nagłówek tabeli; od tej pory zaczynamy zbierać dane
                    started = True
                    continue
                # brak jeszcze nagłówka → pomijamy
                continue

            # 1c) Po wykryciu nagłówka pomijamy powtórzone linie nagłówków:
            if stripped.startswith("Lp") and not stripped.isdigit():
                continue
            if stripped.lower().startswith("nazwa towaru"):
                continue
            if stripped in ["Ilość", "J. miary", "Cena", "Wartość", "netto", "brutto", "Indeks katalogowy", ""]:
                continue

            # 1d) W pozostałych wierszach (czyli faktyczne wiersze z tabeli) dodajemy do all_lines
            all_lines.append(stripped)

    # Krok 2: zlokalizujmy wszystkie indeksy, w których jest Lp (czysta liczba, a następna linia zawiera litery)
    idx_lp = []
    n = len(all_lines)
    for i in range(n - 1):
        if re.fullmatch(r"\d+", all_lines[i]):
            nxt = all_lines[i + 1]
            if re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", nxt) and not nxt.lower().startswith("szt."):
                idx_lp.append(i)

    # Krok 3: zlokalizujmy wszystkie indeksy, w których pojawia się "Kod kres"
    idx_ean = [i for i, ln in enumerate(all_lines) if "Kod kres" in ln]

    # Krok 4: skojarzmy każdy EAN z najbliższym kolejnym Lp (jeśli istnieje)
    ean_map = {}  # {index_Lp: kod_ean}
    for e_idx in idx_ean:
        # wyciągnij tekst po dwukropku jako wartość kodu
        parts = all_lines[e_idx].split(":", maxsplit=1)
        if len(parts) == 2:
            ean_val = parts[1].strip()
            # znajdź pierwsze idx_lp > e_idx
            candidates = [lp_i for lp_i in idx_lp if lp_i > e_idx]
            if candidates:
                target_lp = min(candidates)
                ean_map[target_lp] = ean_val

    # Krok 5: dla każdej pozycji Lp zbudujmy pełny rekord produktu
    products = []
    for lp_idx in idx_lp:
        Lp_val = int(all_lines[lp_idx])

        # 5a) wyciągnięcie nazwy: zbieramy wszystkie wiersze zawierające litery
        #      aż do momentu napotkania ilości (czysta liczba, a następna linia to "szt.")
        name_parts = []
        j = lp_idx + 1
        qty_idx = None

        while j < n:
            ln = all_lines[j]

            # jeżeli linia to czysta liczba i linia dalej to "szt." → to ilość, pora zakończyć zbieranie nazwy
            if re.fullmatch(r"\d+", ln) and (j + 1 < n and all_lines[j + 1].lower() == "szt."):
                qty_idx = j
                break

            # jeżeli linia zawiera litery, nie wygląda na cenę, nie jest VAT-em ani "/" → to fragment nazwy
            if re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", ln):
                # pomijamy formaty wyglądające jak cena: "123,45" lub "1 234,56"
                if not re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", ln) and not ln.startswith("VAT") and ln != "/":
                    name_parts.append(ln)
            j += 1

        # jeżeli nie znaleziono qty_idx, to nie ma kompletnej pozycji → pomijamy
        if qty_idx is None:
            continue

        # 5b) odczyt ilości
        qty = int(all_lines[qty_idx])

        # 5c) scalona nazwa
        Name_val = " ".join(name_parts).strip()

        # 5d) kod EAN, jeśli został przypisany w ean_map do tego lp_idx
        Barcode_val = ean_map.get(lp_idx)

        products.append({
            "Lp": Lp_val,
            "Name": Name_val,
            "Quantity": qty,
            "Barcode": Barcode_val
        })

    # Krok 6: zbudujmy DataFrame i usuńmy wiersze bez Name lub bez Quantity
    df = pd.DataFrame(products)
    df = df.dropna(subset=["Name", "Quantity"]).reset_index(drop=True)
    return df


# ──────────────────────────────────────────────────────────────────────────────

# 1) FileUploader – użytkownik wgrywa PDF
uploaded_file = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if uploaded_file is None:
    st.info("Proszę wgrać plik PDF, aby uruchomić parser.")
    st.stop()

# 2) Wczytanie PDF-a przez PyPDF2
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
