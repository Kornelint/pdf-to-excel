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
       używa starych parserów (układ B, C lub A) – tak było w pierwotnym kodzie.
    3. W przeciwnym razie (lub gdy PyPDF2 nie wyciągnie w ogóle linii) 
       wyciąga tekst przez pdfplumber (nowy sposób) i próbuje wykryć układy:
       - **Układ D**: linie zawierające tylko EAN (13 cyfr) i ilość, np.  
         `5029040012366 Nazwa Produktu 96,00 szt.` lub `5029040012366 96,00 szt.`  
       - **Układ E**: linie zaczynające się od Lp i nazwy, potem ilość, a poniżej „Kod kres.: <EAN>”.  
         (Przykłady: `Gussto wola park.pdf`, `Zamówienie nr ZD 0175_05_25.pdf`.)  
       - **Układ B**: w jednej linii: `<Lp> <EAN(13)> <pełna nazwa> <ilość>,<xx> szt.`  
       - **Układ C**: 13-cyfrowy EAN w osobnej linii, potem Lp, potem nazwa, potem “szt.” i ilość.  
       - **Układ A**: “Kod kres.: <EAN>” w osobnej linii, Lp w osobnej linii, fragmenty nazwy przed i po liczbie.
    4. Wywołuje odpowiedni parser i wyświetla wynik w tabeli:  
       **Lp | Symbol | Ilość** (gdzie “Symbol” to kod EAN).
    5. Umożliwia pobranie danych jako plik Excel.
    """
)


def extract_text_with_pypdf2(pdf_bytes: bytes) -> list[str]:
    """
    Wyciąga wszystkie niepuste linie tekstu przez PyPDF2.
    Jeśli nic nie znajdzie lub wystąpi błąd, zwraca pustą listę.
    """
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(pdf_bytes))
    except Exception:
        return []
    lines: list[str] = []
    for page in reader.pages:
        text = page.extract_text() or ""
        for ln in text.split("\n"):
            stripped = ln.strip()
            if stripped:
                lines.append(stripped)
    return lines


def extract_text_with_pdfplumber(pdf_bytes: bytes) -> list[str]:
    """
    Wyciąga wszystkie niepuste linie tekstu przy pomocy pdfplumber.
    Jeśli nic nie znajdzie lub wystąpi błąd, zwraca pustą listę.
    """
    lines: list[str] = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                for ln in text.split("\n"):
                    stripped = ln.strip()
                    if stripped:
                        lines.append(stripped)
    except Exception:
        return []
    return lines


def parse_layout_d(all_lines: list[str]) -> pd.DataFrame:
    """
    Układ D – linie z EAN i ilością. Symbol = EAN.
    """
    products = []
    pattern = re.compile(r"^(\d{13})(?:\s+.*?)*\s+(\d{1,3}),\d{2}\s+szt", flags=re.IGNORECASE)
    lp_counter = 1
    for ln in all_lines:
        m = pattern.match(ln)
        if m:
            ean = m.group(1)
            qty = int(m.group(2).replace(" ", ""))
            products.append({"Lp": lp_counter, "Symbol": ean, "Ilość": qty})
            lp_counter += 1
    return pd.DataFrame(products)


def parse_layout_e(all_lines: list[str]) -> pd.DataFrame:
    """
    Układ E – linie z Lp i ilością, poniżej "Kod kres.: <EAN>". Symbol = EAN.
    """
    products = []
    pattern_item = re.compile(r"^(\d+)\s+.+?\s+(\d{1,3})\s+szt\.", flags=re.IGNORECASE)
    i = 0
    while i < len(all_lines):
        ln = all_lines[i]
        m = pattern_item.match(ln)
        if m:
            lp = int(m.group(1))
            qty = int(m.group(2))
            ean = ""
            j = i + 1
            while j < len(all_lines):
                nxt = all_lines[j]
                if nxt.lower().startswith("kod kres"):
                    parts = nxt.split(":", 1)
                    if len(parts) == 2:
                        ean = parts[1].strip()
                    j += 1
                    break
                j += 1
            products.append({"Lp": lp, "Symbol": ean, "Ilość": qty})
            i = j
        else:
            i += 1
    return pd.DataFrame(products)


def parse_layout_b(all_lines: list[str]) -> pd.DataFrame:
    """
    Układ B – jedna linia: Lp, EAN, nazwa, ilość. Symbol = EAN.
    """
    products = []
    pattern = re.compile(r"^(\d+)\s+(\d{13})\s+.+?\s+(\d{1,3}),\d{2}\s+szt", flags=re.IGNORECASE)
    for ln in all_lines:
        m = pattern.match(ln)
        if m:
            lp = int(m.group(1))
            ean = m.group(2)
            qty = int(m.group(3).replace(" ", ""))
            products.append({"Lp": lp, "Symbol": ean, "Ilość": qty})
    return pd.DataFrame(products)


def parse_layout_c(all_lines: list[str]) -> pd.DataFrame:
    """
    Układ C – linia z EAN, potem Lp, nazwa, "szt.", ilość. Symbol = EAN.
    """
    idx_lp = []
    for i in range(len(all_lines) - 1):
        if re.fullmatch(r"\d+", all_lines[i]):
            nxt = all_lines[i + 1]
            if re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", nxt):
                idx_lp.append(i)
    idx_ean = [i for i, ln in enumerate(all_lines) if re.fullmatch(r"\d{13}", ln)]
    products = []
    for lp_idx in idx_lp:
        eans_before = [e for e in idx_ean if e < lp_idx]
        ean = all_lines[max(eans_before)] if eans_before else ""
        qty = None
        for j in range(lp_idx + 1, len(all_lines) - 2):
            if all_lines[j].lower() == "szt." and re.fullmatch(r"\d+", all_lines[j + 2]):
                qty = int(all_lines[j + 2])
                break
        if qty is not None:
            lp = int(all_lines[lp_idx])
            products.append({"Lp": lp, "Symbol": ean, "Ilość": qty})
    return pd.DataFrame(products)


def parse_layout_a(all_lines: list[str]) -> pd.DataFrame:
    """
    Układ A – "Kod kres.: <EAN>" w osobnej linii, Lp w osobnej linii, fragmenty nazwy, "szt.", ilość. Symbol = EAN.
    """
    idx_lp = []
    for i in range(len(all_lines) - 1):
        if re.fullmatch(r"\d+", all_lines[i]):
            nxt = all_lines[i + 1]
            if (
                re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", nxt)
                and nxt.lower() != "szt."
                and not re.fullmatch(r"\d{1,3}(?: \d{3})*,\d{2}", nxt)
                and not nxt.lower().startswith("kod kres")
            ):
                idx_lp.append(i)
    idx_ean = [i for i, ln in enumerate(all_lines) if ln.lower().startswith("kod kres")]
    products = []
    for k, lp_idx in enumerate(idx_lp):
        prev_lp = idx_lp[k - 1] if k > 0 else -1
        next_lp = idx_lp[k + 1] if k + 1 < len(idx_lp) else len(all_lines)
        valid_eans = [e for e in idx_ean if prev_lp < e < next_lp]
        ean = ""
        if valid_eans:
            parts = all_lines[max(valid_eans)].split(":", 1)
            if len(parts) == 2:
                ean = parts[1].strip()
        qty = None
        for j in range(lp_idx + 1, next_lp):
            ln = all_lines[j]
            if re.fullmatch(r"\d+", ln) and (j + 1 < next_lp and all_lines[j + 1].lower() == "szt."):
                qty = int(ln)
                break
        if qty is not None:
            lp = int(all_lines[lp_idx])
            products.append({"Lp": lp, "Symbol": ean, "Ilość": qty})
    return pd.DataFrame(products)


# ──────────────────────────────────────────────────────────────────────────────
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

if "Ilość" in df.columns:
    df = df.dropna(subset=["Ilość"]).reset_index(drop=True)

if df.empty:
    st.error("Po parsowaniu nie znaleziono pozycji zamówienia.")
    st.stop()

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
