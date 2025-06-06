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
    1. Próbuję odczytać tekst przez PyPDF2.
    2. Jeśli nie znajdę ani jednej czytelnej linii, wyświetlam informację, że trzeba najpierw wykonać OCR.
    3. Gdy mam listę wierszy (`all_lines`), wykrywam układ:
       - **Układ B**: każda pozycja w jednej linii, np.  
         `1 5029040012366 Nazwa Produktu 96,00 szt.`  
       - **Układ C**: czysty 13-cyfrowy EAN w osobnej linii, potem Lp itp.  
       - **Układ A**: „Kod kres.: <EAN>” w oddzielnej linii, Lp w oddzielnej linii (czysta liczba),  
         fragmenty nazwy przed i po kolumnie cen.  
    4. Wyświetlam tabelę z kolumnami `Lp`, `Name`, `Quantity`, `Barcode` i pozwalam pobrać Excel.
    """
)

# ──────────────────────────────────────────────────────────────────────────────

def extract_text_with_pypdf2(pdf_bytes: bytes) -> list[str]:
    """
    Wyciąga wszystkie niepuste linie tekstu przez PyPDF2.
    Zwraca listę wierszy; jeśli pusta, oznacza brak czytelnego tekstu.
    """
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(pdf_bytes))
    except Exception:
        return []
    lines = []
    for page in reader.pages:
        text = page.extract_text() or ""
        for ln in text.split("\n"):
            stripped = ln.strip()
            if stripped:
                lines.append(stripped)
    return lines


def parse_layout_b(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla rożkładu B – każda pozycja w jednej linii:
      <Lp> <EAN(13)> <pełna nazwa> <ilość>,<xx> szt ...
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


def parse_layout_c(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla układu C – czysty 13-cyfrowy EAN w oddzielnej linii, potem Lp, potem Name, potem "szt." i Quantity.
    """
    # 1) Znajdź indeksy Lp – czysta liczba, pod którą jest tekst
    idx_lp = []
    for i in range(len(all_lines) - 1):
        if re.fullmatch(r"\d+", all_lines[i]):
            nxt = all_lines[i + 1]
            if re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", nxt):
                idx_lp.append(i)

    # 2) Znajdź indeksy czystego EAN (13 cyfr)
    idx_ean = [i for i, ln in enumerate(all_lines) if re.fullmatch(r"\d{13}", ln)]

    products = []
    for idx, lp_idx in enumerate(idx_lp):
        # EAN – maksymalny e < lp_idx
        eans = [e for e in idx_ean if e < lp_idx]
        barcode = all_lines[max(eans)] if eans else None

        # Name = linia zaraz po lp_idx
        Name_val = all_lines[lp_idx + 1] if lp_idx + 1 < len(all_lines) else None

        # Quantity – po znalezieniu "szt.", 2 linie dalej integer
        qty = None
        for j in range(lp_idx + 1, len(all_lines) - 2):
            if all_lines[j].lower() == "szt." and re.fullmatch(r"\d+", all_lines[j + 2]):
                qty = int(all_lines[j + 2])
                break

        if Name_val and qty is not None:
            products.append({
                "Lp": int(all_lines[lp_idx]),
                "Name": Name_val.strip(),
                "Quantity": qty,
                "Barcode": barcode
            })

    return pd.DataFrame(products)


def parse_layout_a(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla układu A – „Kod kres.: <EAN>” w osobnej linii,
    Lp to czysta liczba w osobnej linii, nazwa przed/po kolumnie cen.
    """
    # 1) Indeksy Lp – czysta liczba, pod którą linia z literami
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

    # 2) Indeksy "Kod kres"
    idx_ean = [i for i, ln in enumerate(all_lines) if ln.startswith("Kod kres")]

    products = []
    for idx, lp_idx in enumerate(idx_lp):
        prev_lp = idx_lp[idx - 1] if idx > 0 else -1
        next_lp = idx_lp[idx + 1] if idx + 1 < len(idx_lp) else len(all_lines)

        # a) Barcode – maksymalny e w idx_ean taki, że prev_lp < e < next_lp
        barcode = None
        valid_eans = [e for e in idx_ean if prev_lp < e < next_lp]
        if valid_eans:
            parts = all_lines[max(valid_eans)].split(":", 1)
            if len(parts) == 2:
                barcode = parts[1].strip()

        # b) Zbierz fragmenty nazwy do ilości
        name_parts = []
        qty = None
        qty_idx = None

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

        if qty_idx is None:
            continue

        # c) Po ilości: dopisz fragmenty nazwy aż do "Kod kres"
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

# 1) FileUploader – wgraj PDF
uploaded_file = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if uploaded_file is None:
    st.info("Proszę wgrać plik PDF, aby kontynuować.")
    st.stop()

# 2) Pobierz bajty PDF-a
pdf_bytes = uploaded_file.read()

# 3) Wyciągnij linie tekstu przez PyPDF2
all_lines = extract_text_with_pypdf2(pdf_bytes)

# 4) Jeśli brak linii – komunikat o konieczności OCR i koniec
if not all_lines:
    st.error(
        "Nie udało się wyciągnąć tekstu z tego PDF-a. "
        "Prawdopodobnie jest to skan/obraz. "
        "Aby sparsować, wykonaj OCR (np. narzędziem Tesseract) i spróbuj ponownie."
    )
    st.stop()

# 5) Wykryj układ B (Lp + EAN w jednej linii)
pattern_b = re.compile(r"^\d+\s+\d{13}\s+.+\s+\d{1,3},\d{2}\s+szt", flags=re.IGNORECASE)
is_layout_b = any(pattern_b.match(ln) for ln in all_lines)

# 6) Wykryj układ C (czysty 13-cyfrowy EAN w liniach, ale nie Układ B)
has_pure_ean = any(re.fullmatch(r"\d{13}", ln) for ln in all_lines)
is_layout_c = has_pure_ean and not is_layout_b

# 7) Parsuj w zależności od układu
if is_layout_b:
    df = parse_layout_b(all_lines)
elif is_layout_c:
    df = parse_layout_c(all_lines)
else:
    df = parse_layout_a(all_lines)

# 8) Jeżeli są kolumny "Name" i "Quantity", odfiltrowujemy wiersze z brakami
if "Name" in df.columns and "Quantity" in df.columns:
    df = df.dropna(subset=["Name", "Quantity"]).reset_index(drop=True)

# 9) Wyświetl tabelę w Streamlit
st.subheader("Wyekstrahowane pozycje zamówienia")
st.dataframe(df, use_container_width=True)

# 10) Przygotuj przycisk do pobrania pliku Excel
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
