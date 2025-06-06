# app.py

import streamlit as st
import pandas as pd
import re
import PyPDF2
import io

st.set_page_config(page_title="PDF → Excel", layout="wide")
st.title("PDF → Excel")

st.markdown(
    """
    Wgraj plik PDF ze zamówieniem. Aplikacja:
    1. Najpierw próbuje wyciągnąć tekst przez PyPDF2.
    2. Jeśli w PDF nie ma żadnej czytelnej linii (prawdopodobnie skan/obraz),
       wyświetli komunikat, że PDF wymaga OCR i nie będzie parsowany.
    3. Gdy jest tekst, wykrywa układ:
       - **Układ B**: cała pozycja (Lp, EAN, nazwa, ilość) w jednej linii
         (np. `1 5029040012366 Nazwa Produktu 96,00 szt.`).
       - **Układ A**: EAN w osobnej linii (`Kod kres.: <EAN>`), nazwa może być
         przed i po kolumnie cen, a `Lp` to linia z samą liczbą, pod którą jest
         fragment nazwy.
    4. Wyświetla tabelę z kolumnami `Lp`, `Name`, `Quantity`, `Barcode`.
    5. Pozwala pobrać plik Excel z tymi kolumnami.
    """
)


def extract_text_from_pdf_bytes(pdf_bytes: bytes) -> list[str]:
    """
    Próbujemy wyciągnąć tekst stronami przez PyPDF2. Jeśli nie ma ani jednej
    sensownej linii, zwracamy pustą listę.
    """
    reader = PyPDF2.PdfReader(io.BytesIO(pdf_bytes))
    all_lines = []
    for page in reader.pages:
        text = page.extract_text() or ""
        for ln in text.split("\n"):
            stripped = ln.strip()
            if stripped:
                all_lines.append(stripped)
    return all_lines


def parse_layout_b(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla układu, w którym każda pozycja jest w jednej linii:
      <Lp> <EAN(13)> <pełna nazwa> <ilość>,<xx> szt. <inne kolumny…>
    Wyciąga Lp, Barcode, Name, Quantity.
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
    return pd.DataFrame(products)


def parse_layout_a(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla układu, w którym "Kod kres.: <EAN>" jest w osobnej linii.
    Pozycja Lp to wiersz z samą liczbą, pod którą pojawia się fragment nazwy.
    Nazwa może być przed i po kolumnie cen, a EAN przypisujemy do ostatniego
    wcześniejszego Lp.
    """
    # 1) Zidentyfikuj indeksy Lp: linia czysta-liczba, a pod nią wiersz z literami
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

    # 2) Zidentyfikuj indeksy linii zaczynających się od "Kod kres"
    idx_ean = [i for i, ln in enumerate(all_lines) if ln.startswith("Kod kres")]

    products = []
    for idx, lp_idx in enumerate(idx_lp):
        prev_lp = idx_lp[idx - 1] if idx > 0 else -1
        next_lp = idx_lp[idx + 1] if idx + 1 < len(idx_lp) else len(all_lines)

        # 2a) Barcode: spośród e w idx_ean takich, że prev_lp < e < next_lp, weź ten z największym indeksem
        barcode = None
        valid_eans = [e for e in idx_ean if prev_lp < e < next_lp]
        if valid_eans:
            eidx = max(valid_eans)
            parts = all_lines[eidx].split(":", 1)
            if len(parts) == 2:
                barcode = parts[1].strip()

        # 2b) Name i Quantity:
        name_parts = []
        qty = None
        qty_idx = None

        # Najpierw fragmenty aż do wiersza z ilością (czysta liczba + "szt.")
        for j in range(lp_idx + 1, next_lp):
            ln = all_lines[j]
            if re.fullmatch(r"\d+", ln) and (j + 1 < next_lp and all_lines[j + 1].lower() == "szt."):
                qty_idx = j
                qty = int(ln)
                break
            if (
                re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", ln)
                and not re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", ln)
                and not ln.startswith("VAT")
                and ln != "/"
                and not ln.startswith("ARA")
                and not ln.startswith("KAT")
            ):
                name_parts.append(ln)

        # Jeśli nie znaleziono qty_idx → pomiń
        if qty_idx is None:
            continue

        # Po znalezieniu ilości, zbieramy dodatkowe fragmenty nazwy aż do "Kod kres"
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

    return pd.DataFrame(products)


# ──────────────────────────────────────────────────────────────────────────────

# 1) FileUploader – użytkownik wgrywa PDF
uploaded_file = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if uploaded_file is None:
    st.info("Proszę wgrać plik PDF, aby uruchomić parser.")
    st.stop()

# 2) Pobierz bajty PDF-a
pdf_bytes = uploaded_file.read()

# 3) Spróbuj wyciągnąć tekst przez PyPDF2
all_lines = extract_text_from_pdf_bytes(pdf_bytes)

# 4) Jeśli nie ma żadnej linii (pusty lub obraz), zgłoś, że trzeba zrobić OCR
if not all_lines:
    st.error(
        "Nie udało się wyciągnąć tekstu z PDF-a. "
        "Prawdopodobnie jest to skan/obraz. Aby sparsować, najpierw wykonaj OCR. "
        "Ten parser obsługuje tylko PDF-y z osadzonym „zwykłym” tekstem."
    )
    st.stop()

# 5) Wykryj układ B (Lp + EAN + nazwa + ilość w jednej linii)
pattern_b = re.compile(
    r"^\d+\s+\d{13}\s+.+\s+\d{1,3},\d{2}\s+szt", flags=re.IGNORECASE
)
is_layout_b = any(pattern_b.match(ln) for ln in all_lines)

# 6) Wywołaj odpowiedni parser
if is_layout_b:
    df = parse_layout_b(all_lines)
else:
    df = parse_layout_a(all_lines)

# 7) Odfiltruj wiersze bez nazwy lub ilości
df = df.dropna(subset=["Name", "Quantity"]).reset_index(drop=True)

# 8) Wyświetlenie wyników
st.subheader("Wyekstrahowane pozycje zamówienia")
st.dataframe(df, use_container_width=True)

# 9) Przygotowanie przycisku do pobrania pliku Excel
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
