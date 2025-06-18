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
    3. Jeśli w żaden sposób nie da się wyciągnąć układu D ani E, 
       to używa „nowego” parsera (layout A–E).
    Jeśli nic nie znajdzie lub wystąpi błąd, zwraca pustą listę.
    """
)

# funkcje parsujące różne układy PDF-ów (layout A–E)
def extract_lines_pyppdf(pdf_bytes: bytes) -> list[str]:
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

def extract_lines_pdfplumber(pdf_bytes: bytes) -> list[str]:
    try:
        lines: list[str] = []
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                for ln in page.extract_text().split("\n"):
                    stripped = ln.strip()
                    if stripped:
                        lines.append(stripped)
        return lines
    except Exception:
        return []

def parse_layout_d(all_lines: list[str]) -> pd.DataFrame:
    # ... implementacja parsowania układu D ...
    # przykładowo:
    products = []
    # tu wypełniasz products = [{'Symbol':..., 'Nazwa':..., 'Ilość':..., ...}, ...]
    return pd.DataFrame(products)

def parse_layout_e(all_lines: list[str]) -> pd.DataFrame:
    # ... implementacja parsowania układu E ...
    products = []
    return pd.DataFrame(products)

def parse_layout_b(all_lines: list[str]) -> pd.DataFrame:
    # ... implementacja parsowania układu B ...
    products = []
    return pd.DataFrame(products)

def parse_layout_c(all_lines: list[str]) -> pd.DataFrame:
    # ... implementacja parsowania układu C ...
    products = []
    return pd.DataFrame(products)

def parse_layout_a(all_lines: list[str]) -> pd.DataFrame:
    # ... implementacja parsowania układu A ...
    products = []
    return pd.DataFrame(products)

# wgrywanie pliku
uploaded_file = st.file_uploader("Wybierz plik PDF", type=["pdf"])
if not uploaded_file:
    st.info("Proszę wgrać plik PDF.")
    st.stop()

pdf_bytes = uploaded_file.read()

# ekstrakcja linii
lines_old = extract_lines_pyppdf(pdf_bytes)
layout_text = "\n".join(lines_old)

# sprawdzanie, który parser użyć
lines_new = extract_lines_pdfplumber(pdf_bytes)
# (tutaj logika wykrywająca układ D/E lub B/C/A jak w oryginale)
# poniżej przykładowy fragment:
is_layout_d_new = bool(re.search(r"Twój_wzorzec_D", layout_text))
has_kod_kres_new = any(re.fullmatch(r"\d{13}", ln) for ln in lines_new)
is_layout_e_new = any(re.search(r"Twój_wzorzec_E", ln) for ln in lines_new) and has_kod_kres_new
is_layout_b_new = any(re.compile(r"^\d+\s+\d{13}\s+.+\s\d{2}\s+szt", flags=re.IGNORECASE).match(ln) for ln in lines_new)
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

# czyszczenie kolumny Ilość
if "Ilość" in df.columns:
    df = df.dropna(subset=["Ilość"]).reset_index(drop=True)

# jeśli puste, kończymy
if df.empty:
    st.error("Po parsowaniu nie znaleziono pozycji zamówienia.")
    st.stop()

# --- TU WSTAWIONO NOWY KOD LICZNIKA EAN ---

total_eans = df.shape[0]
unique_eans = df["Symbol"].nunique()
total_qty = int(df["Ilość"].sum())

st.markdown(
    f"**Znaleziono w sumie:** {total_eans} pozycji z kodami EAN  \n"
    f"**Unikalnych kodów EAN:** {unique_eans}  \n"
    f"**Łączna suma ilości:** {total_qty}"
)

# --- KONIEC DODATKU ---

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
