import streamlit as st
import pandas as pd
import re
import io
import PyPDF2
import pdfplumber

st.set_page_config(page_title="PDF → Excel", layout="wide")
st.title("PDF → Excel")

st.markdown(
    """
    Wgraj plik PDF ze zamówieniem. Aplikacja:
    1. Próbuje wyciągnąć tekst przez PyPDF2 (stare „trudniejsze” PDF-y).
    2. Jeśli w wyciągniętym przez PyPDF2 tekście nie występują układy D ani E, 
       używa starych parserów (układ B, C lub A).
    3. W przeciwnym razie (lub gdy PyPDF2 nie wyciągnie w ogóle linii) wyciąga tekst przez pdfplumber
       i próbuje wykryć układy D, E, B, C, A.
    4. Wywołuje odpowiedni parser i wyświetla wynik w tabeli.
    5. Umożliwia pobranie danych jako plik Excel.
    """
)

def extract_text_with_pypdf2(pdf_bytes: bytes) -> list[str]:
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(pdf_bytes))
    except Exception:
        return []
    lines: list[str] = []
    for page in reader.pages:
        text = page.extract_text() or ""
        for ln in text.split("\n"):
            if (stripped := ln.strip()):
                lines.append(stripped)
    return lines

def extract_text_with_pdfplumber(pdf_bytes: bytes) -> list[str]:
    try:
        lines: list[str] = []
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                for ln in (page.extract_text() or "").split("\n"):
                    if (stripped := ln.strip()):
                        lines.append(stripped)
        return lines
    except Exception:
        return []

def parse_layout_d(all_lines: list[str]) -> pd.DataFrame:
    products = []
    pattern = re.compile(r"^(\d{13})(?:\s+.*?)*\s+(\d{1,3}),\d{2}\s+szt", flags=re.IGNORECASE)
    lp_counter = 1
    for ln in all_lines:
        if m := pattern.match(ln):
            ean = m.group(1)
            qty = int(m.group(2).replace(" ", ""))
            products.append({"Lp": lp_counter, "Symbol": ean, "Ilość": qty})
            lp_counter += 1
    return pd.DataFrame(products)

def parse_layout_e(all_lines: list[str]) -> pd.DataFrame:
    products = []
    pattern_item = re.compile(r"^(\d+)\s+.+?\s+(\d{1,3})\s+szt\.", flags=re.IGNORECASE)
    i = 0
    while i < len(all_lines):
        if m := pattern_item.match(all_lines[i]):
            lp = int(m.group(1))
            qty = int(m.group(2))
            # szukamy kodu kreskowego poniżej
            ean = ""
            j = i + 1
            while j < len(all_lines):
                if all_lines[j].lower().startswith("kod kres"):
                    parts = all_lines[j].split(":", 1)
                    if len(parts) == 2:
                        ean = parts[1].strip()
                    break
                j += 1
            products.append({"Lp": lp, "Symbol": ean, "Ilość": qty})
            i = j + 1
        else:
            i += 1
    return pd.DataFrame(products)

def parse_layout_b(all_lines: list[str]) -> pd.DataFrame:
    products = []
    pattern = re.compile(r"^(\d+)\s+(\d{13})\s+.+?\s+(\d{1,3}),\d{2}\s+szt", flags=re.IGNORECASE)
    for ln in all_lines:
        if m := pattern.match(ln):
            products.append({
                "Lp": int(m.group(1)),
                "Symbol": m.group(2),
                "Ilość": int(m.group(3).replace(" ", ""))
            })
    return pd.DataFrame(products)

def parse_layout_c(all_lines: list[str]) -> pd.DataFrame:
    idx_lp = [
        i for i in range(len(all_lines) - 1)
        if re.fullmatch(r"\d+", all_lines[i])
           and re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", all_lines[i+1])
           and not all_lines[i+1].lower().startswith("kod kres")
    ]
    idx_ean = [i for i, ln in enumerate(all_lines) if re.fullmatch(r"\d{13}", ln)]
    products = []
    for lp_idx in idx_lp:
        # znajdź najbliższy kod kres. przed Lp
        before = [e for e in idx_ean if e < lp_idx]
        ean = all_lines[max(before)] if before else ""
        # znajdź ilość za Lp
        qty = None
        for j in range(lp_idx + 1, len(all_lines) - 1):
            if all_lines[j].lower() == "szt." and re.fullmatch(r"\d+", all_lines[j+1]):
                qty = int(all_lines[j+1]); break
        if qty is not None:
            products.append({"Lp": int(all_lines[lp_idx]), "Symbol": ean, "Ilość": qty})
    return pd.DataFrame(products)

def parse_layout_a(all_lines: list[str]) -> pd.DataFrame:
    idx_lp = [
        i for i in range(len(all_lines) - 1)
        if re.fullmatch(r"\d+", all_lines[i])
           and re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", all_lines[i+1])
           and not all_lines[i+1].lower().startswith("kod kres")
    ]
    idx_ean = [i for i, ln in enumerate(all_lines) if ln.lower().startswith("kod kres")]
    products = []
    for k, lp_idx in enumerate(idx_lp):
        prev_lp = idx_lp[k-1] if k > 0 else -1
        next_lp = idx_lp[k+1] if k+1 < len(idx_lp) else len(all_lines)
        valid = [e for e in idx_ean if prev_lp < e < next_lp]
        ean = ""
        if valid:
            parts = all_lines[max(valid)].split(":", 1)
            if len(parts) == 2:
                ean = parts[1].strip()
        # szukamy ilości
        qty = None
        for j in range(lp_idx + 1, next_lp):
            if re.fullmatch(r"\d+", all_lines[j]) and all_lines[j+1].lower() == "szt.":
                qty = int(all_lines[j]); break
        if qty is not None:
            products.append({"Lp": int(all_lines[lp_idx]), "Symbol": ean, "Ilość": qty})
    return pd.DataFrame(products)

# ────────────────────────────────────────────────────────────────────────────
# 3) GŁÓWNA LOGIKA: WCZYTANIE PLIKÓW I WYBOR PARSERA
uploaded_file = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if uploaded_file is None:
    st.info("Proszę wgrać plik PDF, aby kontynuować.")
    st.stop()

pdf_bytes = uploaded_file.read()

# 3.2) PyPDF2
lines_py = extract_text_with_pypdf2(pdf_bytes)
pattern_d = re.compile(r"^\d{13}(?:\s+.*?)*\s+\d{1,3},\d{2}\s+szt", flags=re.IGNORECASE)
is_layout_d_py = any(pattern_d.match(ln) for ln in lines_py)
pattern_e = re.compile(r"^\d+\s+.+?\s+\d{1,3}\s+szt\.", flags=re.IGNORECASE)
has_kod_kres_py = any(ln.lower().startswith("kod kres") for ln in lines_py)
is_layout_e_py = any(pattern_e.match(ln) for ln in lines_py) and has_kod_kres_py

df = pd.DataFrame()
if lines_py and not is_layout_d_py and not is_layout_e_py:
    pattern_b = re.compile(r"^\d+\s+\d{13}\s+.+\s+\d{1,3},\d{2}\s+szt", flags=re.IGNORECASE)
    is_layout_b_py = any(pattern_b.match(ln) for ln in lines_py)
    has_pure_ean_py = any(re.fullmatch(r"\d{13}", ln) for ln in lines_py)
    is_layout_c_py = has_pure_ean_py and not is_layout_b_py

    if is_layout_b_py:
        df = parse_layout_b(lines_py)
    elif is_layout_c_py:
        df = parse_layout_c(lines_py)
    else:
        df = parse_layout_a(lines_py)

# jeśli nic nie wyszło – pdfplumber
if df.empty:
    lines_new = extract_text_with_pdfplumber(pdf_bytes)
    if not lines_new:
        st.error("Nie udało się wyciągnąć tekstu z tego PDF-a. Wykonaj OCR i wgraj ponownie.")
        st.stop()

    is_layout_d_new = any(pattern_d.match(ln) for ln in lines_new)
    has_kod_kres_new = any(ln.lower().startswith("kod kres") for ln in lines_new)
    is_layout_e_new = any(pattern_e.match(ln) for ln in lines_new) and has_kod_kres_new
    is_layout_b_new = any(re.compile(r"^\d+\s+\d{13}\s+.+\s+\d{1,3},\d{2}\s+szt", flags=re.IGNORECASE).match(ln) for ln in lines_new)
    has_pure_ean_new = any(re.fullmatch(r"\d{13}", ln) for ln in lines_new)
    is_layout_c_new = has_pure_ean_new and not is_layout_b_new

    if is_layout_d_new:
        df = parse_layout_d(lines_new)
    elif is_layout_e_new:
        df = parse_layout_e(lines_new)
    elif is_layout_b_new:
        df = parse_layout_b(lines_new)
    elif is_layout_c_new:
        df = parse_layout_c(lines_new)
    else:
        df = parse_layout_a(lines_new)

# usuwamy puste ilości
if "Ilość" in df.columns:
    df = df.dropna(subset=["Ilość"]).reset_index(drop=True)

if df.empty:
    st.error("Po parsowaniu nie znaleziono pozycji zamówienia.")
    st.stop()

# ─── TU DODAJEMY STATYSTYKI ────────────────────────────────────────────
total_eans = df.shape[0]
unique_eans = df["Symbol"].nunique()
total_qty = int(df["Ilość"].sum())

st.markdown(
    f"**Znaleziono w sumie:** {total_eans} pozycji z kodami EAN  \n"
    f"**Unikalnych kodów EAN:** {unique_eans}  \n"
    f"**Łączna suma ilości:** {total_qty}"
)
# ───────────────────────────────────────────────────────────────────────

st.subheader("Wyekstrahowane pozycje zamówienia")
st.dataframe(df, use_container_width=True)

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
