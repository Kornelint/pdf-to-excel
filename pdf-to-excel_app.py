# app.py

import streamlit as st
import pandas as pd
import re
import PyPDF2
import io
from pdf2image import convert_from_bytes
import pytesseract

st.set_page_config(page_title="PDF → Excel", layout="wide")
st.title("PDF → Excel (w tym skany/OCR)")

st.markdown(
    """
    Wgraj plik PDF ze zamówieniem. Aplikacja:
    1. Próbuje wyciągnąć tekst bezpośrednio (PyPDF2).  
    2. Jeśli nie uda się odczytać żadnej treści (np. PDF to skan), wykonuje OCR (pytesseract) na stronach PDF.  
    3. Na uzyskanym tekście wykrywa układ:
       - **Układ B**: cała pozycja (Lp, EAN, nazwa, ilość) w jednej linii, 
         np. `1 5029040012366 Nazwa Produktu 96,00 szt. …`.  
       - **Układ A**: Lp i nazwa mogą być w różnych wierszach, a „Kod kres.: <EAN>” jest w osobnej linii.  
    4. W rezultacie wyświetla tabelę z kolumnami `Lp`, `Name`, `Quantity`, `Barcode`.  
    5. Umożliwia pobranie wyniku jako plik Excel.
    """
)

# ──────────────────────────────────────────────────────────────────────────────

def extract_text_from_pdf(uploaded_bytes: bytes) -> list[str]:
    """
    Próbuje zbierać surowy tekst stronami przez PyPDF2; jeśli nie znajdzie żadnej czytelnej linii,
    przechodzi do OCR (pdf2image + pytesseract).
    Zwraca listę linii tekstu.
    """
    # 1) Spróbuj PyPDF2
    reader = PyPDF2.PdfReader(io.BytesIO(uploaded_bytes))
    all_lines = []
    for page in reader.pages:
        text = page.extract_text() or ""
        for ln in text.split("\n"):
            stripped = ln.strip()
            if stripped:
                all_lines.append(stripped)
    # Jeśli wyciągnięto co najmniej jedną sensowną linię, zwracamy:
    if len(all_lines) > 0:
        return all_lines

    # 2) W przeciwnym razie: OCR (pdf2image + pytesseract)
    all_lines = []
    # Konwertuj strony PDF na obrazy:
    images = convert_from_bytes(uploaded_bytes)
    for img in images:
        # pytesseract OCR (domyślnie język polski, jeśli masz zainstalowany 'pol'):
        text = pytesseract.image_to_string(img, lang="pol")
        for ln in text.split("\n"):
            stripped = ln.strip()
            if stripped:
                all_lines.append(stripped)
    return all_lines


def parse_layout_b(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla układu, w którym każda pozycja jest w jednej linii:
      <Lp> <EAN(13)> <pełna nazwa> <ilość>,<xx> szt. <…inne kolumny…>
    Wyciągamy: Lp, Barcode, Name, Quantity.
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
    Pozycje Lp to linie czystych liczb, pod którymi jest fragment nazwy. Nazwa może być 
    przed i po kolumnach cen, a EAN przypisujemy do ostatniego wcześniejszego Lp.
    """
    # 1) Zidentyfikuj wszystkie indeksy Lp: linia czysta-liczba, a pod nią wiersz z literami
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

    # 2) Zidentyfikuj wszystkie indeksy EAN (linia zaczynająca się od "Kod kres")
    idx_ean = [i for i, ln in enumerate(all_lines) if ln.startswith("Kod kres")]

    products = []
    for idx, lp_idx in enumerate(idx_lp):
        prev_lp = idx_lp[idx - 1] if idx > 0 else -1
        next_lp = idx_lp[idx + 1] if idx + 1 < len(idx_lp) else len(all_lines)

        # 2a) Barcode: spośród e w idx_ean takich, że prev_lp < e < next_lp, weź największy
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
            # jeżeli linia to czysta liczba i linia poniżej to "szt." → to ilość
            if re.fullmatch(r"\d+", ln) and (j + 1 < next_lp and all_lines[j + 1].lower() == "szt."):
                qty_idx = j
                qty = int(ln)
                break
            # w przeciwnym razie, jeśli ln zawiera litery i nie wygląda jak cena/VAT/"ARA"/"KAT"/"/" → fragment nazwy
            if (
                re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", ln)
                and not re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", ln)
                and not ln.startswith("VAT")
                and ln != "/"
                and not ln.startswith("ARA")
                and not ln.startswith("KAT")
            ):
                name_parts.append(ln)

        # Jeśli nie znaleziono qty_idx → pomiń tę pozycję
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


def parse_pdf_generic(reader: PyPDF2.PdfReader) -> pd.DataFrame:
    """
    Główny parser:
      1) Pobiera wszystkie wiersze z PDF (text lub OCR) → all_lines
      2) Jeśli przynajmniej jeden wiersz pasuje do wzorca Układu B, 
         wywołuje parse_layout_b(all_lines).  
      3) W przeciwnym razie wywołuje parse_layout_a(all_lines).
    """
    all_lines = []
    started = False

    # 1) Pobieramy linie przez extract_text() i OCR, ale tu już mamy 'all_lines'
    #    – zostaniemy wywołani z tą listą (poza tą funkcją).
    # Nie potrzebujemy tu dodatkowo czytać tekstu, bo robimy to wyżej.

    # Ta funkcja zakłada, że otrzyma `all_lines` jako argument (z zewnątrz).  
    # W rzeczywistości będziemy ją wywoływali w miejscu, gdzie mamy `all_lines`.
    raise RuntimeError("parse_pdf_generic() nie powinno być wywoływane bezpośrednio.")


# ──────────────────────────────────────────────────────────────────────────────

# 1) FileUploader – użytkownik wgrywa PDF
uploaded_file = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if uploaded_file is None:
    st.info("Proszę wgrać plik PDF, aby uruchomić parser.")
    st.stop()

# 2) Pobierz bajty wgrywanego PDF-a
pdf_bytes = uploaded_file.read()

# 3) Wyciągnij linie tekstu (PyPDF2 lub OCR)
all_lines = extract_text_from_pdf(pdf_bytes)

# 4) Teraz, gdy mamy all_lines, wykrywamy, który układ:
pattern_b = re.compile(r"^\d+\s+\d{13}\s+.+\s+\d{1,3},\d{2}\s+szt", flags=re.IGNORECASE)
is_layout_b = any(pattern_b.match(ln) for ln in all_lines)

if is_layout_b:
    df = parse_layout_b(all_lines)
else:
    df = parse_layout_a(all_lines)

# 5) Usuń ewentualne wiersze bez nazwy lub ilości (żeby nie było pustych)
df = df.dropna(subset=["Name", "Quantity"]).reset_index(drop=True)

# 6) Wyświetlenie wyników
st.subheader("Wyekstrahowane pozycje zamówienia")
st.dataframe(df, use_container_width=True)

# 7) Przygotowanie przycisku Excel
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
