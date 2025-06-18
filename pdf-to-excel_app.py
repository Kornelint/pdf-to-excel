import streamlit as st
import pandas as pd
import re
import io
import pdfplumber

st.set_page_config(page_title="PDF → Excel", layout="wide")
st.title("PDF → Excel")

st.markdown(
    """
    Wgraj plik PDF ze zamówieniem. Aplikacja:
    1. Wyciąga wszystkie linie przez pdfplumber.
    2. Usuwa stopki/numerację stron.
    3. Wykrywa layout (D, E, B, C lub A) i parsuje odpowiednio.
    4. Wyświetla tabelę wynikową oraz statystyki.
    5. Umożliwia pobranie wyniku jako plik Excel.
    """
)


def extract_text_with_pdfplumber(pdf_bytes: bytes) -> list[str]:
    """
    Wyciąga wszystkie niepuste linie tekstu przy pomocy pdfplumber.
    Jeśli nic nie znajdzie lub wystąpi błąd, zwraca pustą listę.
    """
    try:
        lines: list[str] = []
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                for ln in text.split("\n"):
                    stripped = ln.strip()
                    if stripped:
                        lines.append(stripped)
        return lines
    except Exception:
        return []


def parse_layout_d(all_lines: list[str]) -> pd.DataFrame:
    products = []
    pattern = re.compile(r"^(\d{13})(?:\s+.*?)*\s+(\d{1,3}),\d{2}\s+szt", flags=re.IGNORECASE)
    lp = 1
    for ln in all_lines:
        if m := pattern.match(ln):
            ean = m.group(1)
            qty = int(m.group(2).replace(" ", ""))
            products.append({"Lp": lp, "Symbol": ean, "Ilość": qty})
            lp += 1
    return pd.DataFrame(products)


def parse_layout_e(all_lines: list[str]) -> pd.DataFrame:
    products = []
    pattern_item = re.compile(r"^(\d+)\s+.+?\s+(\d{1,3})\s+szt\.", flags=re.IGNORECASE)
    i = 0
    while i < len(all_lines):
        if m := pattern_item.match(all_lines[i]):
            lp = int(m.group(1))
            qty = int(m.group(2))
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
    # znajdź indeksy Lp i czystych kodów EAN
    idx_lp = [
        i for i in range(len(all_lines)-1)
        if re.fullmatch(r"\d+", all_lines[i])
           and re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", all_lines[i+1])
           and not all_lines[i+1].lower().startswith("kod kres")
    ]
    idx_ean = [i for i, ln in enumerate(all_lines) if re.fullmatch(r"\d{13}", ln)]
    products = []
    for lp_idx in idx_lp:
        prev_lp = max([e for e in idx_lp if e < lp_idx], default=-1)
        next_lp = min([e for e in idx_lp if e > lp_idx], default=len(all_lines))
        valid = [e for e in idx_ean if prev_lp < e < next_lp]
        ean = all_lines[max(valid)] if valid else ""
        qty = None
        for j in range(lp_idx+1, next_lp):
            if re.fullmatch(r"\d+", all_lines[j]) and j+1 < next_lp and all_lines[j+1].lower()=="szt.":
                qty = int(all_lines[j]); break
        if qty is not None:
            products.append({"Lp": int(all_lines[lp_idx]), "Symbol": ean, "Ilość": qty})
    return pd.DataFrame(products)


def parse_layout_a(all_lines: list[str]) -> pd.DataFrame:
    idx_lp = [
        i for i in range(len(all_lines)-1)
        if re.fullmatch(r"\d+", all_lines[i])
           and re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", all_lines[i+1])
           and not all_lines[i+1].lower().startswith("kod kres")
    ]
    idx_kod = [i for i, ln in enumerate(all_lines) if ln.lower().startswith("kod kres")]
    products = []
    for k, lp_idx in enumerate(idx_lp):
        prev_lp = idx_lp[k-1] if k>0 else -1
        next_lp = idx_lp[k+1] if k+1<len(idx_lp) else len(all_lines)
        valid = [e for e in idx_kod if prev_lp < e < next_lp]
        ean = ""
        if valid:
            parts = all_lines[max(valid)].split(":",1)
            if len(parts)==2: ean = parts[1].strip()
        qty = None
        for j in range(lp_idx+1, next_lp):
            if re.fullmatch(r"\d+", all_lines[j]) and j+1<next_lp and all_lines[j+1].lower()=="szt.":
                qty = int(all_lines[j]); break
        if qty is not None:
            products.append({"Lp": int(all_lines[lp_idx]), "Symbol": ean, "Ilość": qty})
    return pd.DataFrame(products)


# ────────────────────────────────────────────────────────────────────────────
# GŁÓWNA LOGIKA: zawsze pdfplumber i filtrowanie stron

uploaded_file = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if not uploaded_file:
    st.info("Proszę wgrać plik PDF, aby kontynuować.")
    st.stop()

pdf_bytes = uploaded_file.read()

# 1) wyciągamy wszystkie linie z pdfplumber
lines_all = extract_text_with_pdfplumber(pdf_bytes)

# 2) usuwamy stopki/numerację stron
lines_all = [
    ln for ln in lines_all
    if not ln.startswith("/")      # np. "/ Wydrukowano z programu…"
       and "Strona" not in ln      # linie typu "Strona 1/2"
]

# 3) wykrywamy layout
pattern_d = re.compile(r"^(\d{13})(?:\s+.*?)*\s+(\d{1,3}),\d{2}\s+szt", flags=re.IGNORECASE)
pattern_e = re.compile(r"^(\d+)\s+.+?\s+(\d{1,3})\s+szt\.", flags=re.IGNORECASE)
is_d = any(pattern_d.match(ln) for ln in lines_all)
has_kres = any(ln.lower().startswith("kod kres") for ln in lines_all)
is_e = any(pattern_e.match(ln) for ln in lines_all) and has_kres
is_b = any(re.compile(r"^\d+\s+\d{13}\s+.+?\s+\d{1,3},\d{2}\s+szt", flags=re.IGNORECASE).match(ln)
           for ln in lines_all)
has_plain_ean = any(re.fullmatch(r"\d{13}", ln) for ln in lines_all)
is_c = has_plain_ean and not is_b

# 4) parsujemy odpowiednią funkcją
if is_d:
    df = parse_layout_d(lines_all)
elif is_e:
    df = parse_layout_e(lines_all)
elif is_b:
    df = parse_layout_b(lines_all)
elif is_c:
    df = parse_layout_c(lines_all)
else:
    df = parse_layout_a(lines_all)

# 5) drop pustych ilości i sprawdzenie
if "Ilość" in df.columns:
    df = df.dropna(subset=["Ilość"]).reset_index(drop=True)

if df.empty:
    st.error("Po parsowaniu nie znaleziono pozycji zamówienia.")
    st.stop()

# 6) statystyki EAN-ów
total_eans = df.shape[0]
unique_eans = df["Symbol"].nunique()
total_qty = int(df["Ilość"].sum())

st.markdown(
    f"**Znaleziono w sumie:** {total_eans} pozycji z kodami EAN  \n"
    f"**Unikalnych kodów EAN:** {unique_eans}  \n"
    f"**Łączna suma ilości:** {total_qty}"
)

# 7) wyświetlenie tabeli i eksport do Excela
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
