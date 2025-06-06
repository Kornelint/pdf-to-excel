# app.py

import streamlit as st
import pandas as pd
import re
import PyPDF2
import io

# Spróbuj zaimportować pdfplumber – jeśli jest zainstalowane, użyjemy go jako fallback do wydobycia tekstu
try:
    import pdfplumber
    _HAS_PDFPLUMBER = True
except ImportError:
    _HAS_PDFPLUMBER = False

st.set_page_config(page_title="PDF → Excel", layout="wide")
st.title("PDF → Excel")

if not _HAS_PDFPLUMBER:
    st.warning(
        "Uwaga: w środowisku nie zainstalowano biblioteki `pdfplumber`. "
        "Jeśli PyPDF2 nie będzie w stanie wydobyć tekstu z niektórych PDF-ów (np. „Wydruk.pdf”), "
        "zainstaluj `pdfplumber` przez `pip install pdfplumber`."
    )

st.markdown(
    """
    Wgraj plik PDF ze zamówieniem. Aplikacja:
    1. Wyciąga tekst przez PyPDF2.  
    2. Jeśli PyPDF2 nie zwróci sensownych linii (np. „gibberish”), a `pdfplumber` jest dostępny, 
       próbuje ponownie z `pdfplumber`.  
    3. Gdy mamy listę wierszy (`all_lines`), wykrywamy układ:
       - **Układ B**: jedna pozycja w jednym wierszu, np.  
         `1 5029040012366 Nazwa Produktu 96,00 szt.`  
       - **Układ C**: czysty 13-cyfrowy EAN w osobnej linii, potem Lp, potem nazwę, potem „szt.” i ilość.  
       - **Układ A**: wiersz „Kod kres.: <EAN>”, Lp w osobnej linii, fragmenty nazwy przed i po kolumnie ceny/ilości.  
    4. W zależności od wykrytego układu wywołujemy odpowiedni parser (A, B lub C).  
    5. Wyświetlamy tabelę z kolumnami `Lp`, `Name`, `Quantity`, `Barcode` i umożliwiamy pobranie pliku Excel.
    """
)


def extract_text(pdf_bytes: bytes) -> list[str]:
    """
    Najpierw próbuje wydobyć tekst przy pomocy PyPDF2.
    Jeśli wynik nie zawiera przynajmniej jednej linii pasującej do wzorca (Lp/EAN) 
    i jeśli pdfplumber jest zainstalowane, próbuje ponownie z pdfplumber.
    Zwraca listę niepustych, obciętych linii.
    """
    # 1) PyPDF2
    lines = []
    try:
        reader = PyPDF2.PdfReader(io.BytesIO(pdf_bytes))
        for page in reader.pages:
            text = page.extract_text() or ""
            for ln in text.split("\n"):
                stripped = ln.strip()
                if stripped:
                    lines.append(stripped)
    except Exception:
        lines = []

    # Sprawdźmy, czy PyPDF2 zwróciło przynajmniej jedną linię zaczynającą się od liczby lub „Kod kres.”
    has_layout_indicator = any(
        re.match(r"^\d+", ln) or ln.lower().startswith("kod kres") 
        for ln in lines
    )
    if has_layout_indicator and lines:
        return lines

    # 2) Fallback na pdfplumber (jeśli dostępne)
    if _HAS_PDFPLUMBER:
        pl_lines = []
        try:
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                for page in pdf.pages:
                    text = page.extract_text() or ""
                    for ln in text.split("\n"):
                        stripped = ln.strip()
                        if stripped:
                            pl_lines.append(stripped)
        except Exception:
            pl_lines = []

        if pl_lines:
            return pl_lines

    # Jeśli wszystko zawiedzie, zwracamy to, co mamy (może być puste)
    return lines


def parse_layout_b(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla układu B – każda pozycja w jednej linii:
      <Lp> <EAN(13)> <pełna nazwa> <ilość>,<xx> szt …
    Regex: ^(\d+)\s+(\d{13})\s+(.+?)\s+([\d,]+)\s+szt
    """
    products = []
    pattern = re.compile(
        r"^(\d+)\s+(\d{13})\s+(.+?)\s+([\d,]+)\s+szt", 
        flags=re.IGNORECASE
    )
    for ln in all_lines:
        m = pattern.match(ln)
        if m:
            Lp_val = int(m.group(1))
            Barcode_val = m.group(2)
            Name_val = m.group(3).strip()
            qty_str = m.group(4).replace(",", ".")
            try:
                Quantity_val = float(qty_str)
            except ValueError:
                Quantity_val = 0.0
            products.append({
                "Lp": Lp_val,
                "Name": Name_val,
                "Quantity": Quantity_val,
                "Barcode": Barcode_val
            })
    return pd.DataFrame(products)


def parse_layout_c(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla układu C – czysty 13-cyfrowy EAN w osobnej linii, potem Lp, potem Name, 
    potem "szt." i Quantity.
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
        # Przypisz EAN – ostatni indeks < lp_idx
        eans = [e for e in idx_ean if e < lp_idx]
        barcode = all_lines[max(eans)] if eans else None

        Name_val = all_lines[lp_idx + 1] if lp_idx + 1 < len(all_lines) else None

        qty = None
        for j in range(lp_idx + 1, len(all_lines) - 2):
            if all_lines[j].lower() == "szt." and re.fullmatch(r"[\d,]+", all_lines[j + 1]):
                qty_str = all_lines[j + 1].replace(",", ".")
                try:
                    qty = float(qty_str)
                except ValueError:
                    qty = None
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
    Lp to czysta liczba w osobnej linii, fragmenty nazwy przed i po kolumnie cen/ilości.
    """
    idx_lp = []
    for i in range(len(all_lines) - 1):
        if re.fullmatch(r"\d+", all_lines[i]):
            nxt = all_lines[i + 1]
            if (
                re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", nxt)
                and nxt.lower() != "szt."
                and not re.fullmatch(r"[\d,]+\s*szt", nxt, re.IGNORECASE)
                and not nxt.lower().startswith("kod kres")
            ):
                idx_lp.append(i)

    idx_ean = [i for i, ln in enumerate(all_lines) if ln.lower().startswith("kod kres")]

    products = []
    for idx, lp_idx in enumerate(idx_lp):
        prev_lp = idx_lp[idx - 1] if idx > 0 else -1
        next_lp = idx_lp[idx + 1] if idx + 1 < len(idx_lp) else len(all_lines)

        # Spośród linii „Kod kres…” w przedziale (prev_lp, next_lp) wybierz ostatnią
        valid_eans = [e for e in idx_ean if prev_lp < e < next_lp]
        barcode = None
        if valid_eans:
            parts = all_lines[max(valid_eans)].split(":", 1)
            if len(parts) == 2:
                barcode = parts[1].strip()

        # Zbierz fragmenty nazwy aż do wiersza z ilością
        name_parts = []
        qty = None
        qty_idx = None

        for j in range(lp_idx + 1, next_lp):
            ln = all_lines[j]
            if re.fullmatch(r"[\d,]+\s*szt", ln, re.IGNORECASE):
                # np. "96,00 szt"
                m_qty = re.search(r"([\d,]+)\s*szt", ln, re.IGNORECASE)
                if m_qty:
                    qty_str = m_qty.group(1).replace(",", ".")
                    try:
                        qty = float(qty_str)
                    except ValueError:
                        qty = None
                qty_idx = j
                break
            if (
                re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", ln)
                and not ln.lower().startswith("vat")
                and ln != "/"
                and not ln.lower().startswith("kod kres")
            ):
                name_parts.append(ln)

        if qty_idx is None:
            continue

        # Dodatkowe fragmenty nazwy po qty_idx, aż do next_lp
        for k in range(qty_idx + 1, next_lp):
            ln2 = all_lines[k]
            if ln2.lower().startswith("kod kres"):
                break
            if (
                re.search(r"[A-Za-zĄĆĘŁŃÓŚŹŻąćęłńóśźż]", ln2)
                and not ln2.lower().startswith("vat")
                and ln2 != "/"
            ):
                name_parts.append(ln2)

        Name_val = " ".join(name_parts).strip()
        if Name_val and qty is not None:
            products.append({
                "Lp": int(all_lines[lp_idx]),
                "Name": Name_val,
                "Quantity": qty,
                "Barcode": barcode
            })

    return pd.DataFrame(products)


# ──────────────────────────────────────────────────────────────────────────────

# 1) Wgraj PDF
uploaded_file = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if uploaded_file is None:
    st.info("Proszę wgrać plik PDF, aby kontynuować.")
    st.stop()

pdf_bytes = uploaded_file.read()

# 2) Wyciągnij linie tekstu (PyPDF2 + opcjonalnie pdfplumber)
all_lines = extract_text(pdf_bytes)

# 3) Jeśli brak linii – komunikat o konieczności OCR
if not all_lines:
    st.error(
        "Nie udało się wyciągnąć tekstu z tego PDF-a. "
        "Prawdopodobnie wymaga OCR lub ma niestandardową warstwę czcionek. "
        "Jeśli masz `pdfplumber`, upewnij się, że jest zainstalowane."
    )
    st.stop()

# 4) Wykryj układ B (jedna pozycja w jednej linii)
pattern_b = re.compile(r"^\d+\s+\d{13}\s+.+\s+[\d,]+\s+szt", flags=re.IGNORECASE)
is_layout_b = any(pattern_b.match(ln) for ln in all_lines)

# 5) Wykryj układ C (czysty 13-cyfrowy EAN, ale nie układ B)
has_pure_ean = any(re.fullmatch(r"\d{13}", ln) for ln in all_lines)
is_layout_c = has_pure_ean and not is_layout_b

# 6) Parsuj w zależności od układu
if is_layout_b:
    df = parse_layout_b(all_lines)
elif is_layout_c:
    df = parse_layout_c(all_lines)
else:
    df = parse_layout_a(all_lines)

# 7) Odfiltruj wiersze bez nazwy lub ilości (jeśli te kolumny istnieją)
if "Name" in df.columns and "Quantity" in df.columns:
    df = df.dropna(subset=["Name", "Quantity"]).reset_index(drop=True)

# 8) Jeśli DataFrame pusty → komunikat
if df.empty:
    st.warning(
        "Nie znaleziono żadnych pozycji w pliku PDF. "
        "Sprawdź, czy dokument jest zgodny z jednym z obsługiwanych układów:\n"
        "- Układ B: `Lp EAN Nazwa ilość,xx szt`\n"
        "- Układ C: linia z 13 cyframi (EAN), potem Lp, Nazwa, `szt.` i ilość\n"
        "- Układ A: linia `Kod kres.: EAN`, potem Lp, fragmenty nazwy, `ilość szt.`"
    )
    st.stop()

# 9) Wyświetlenie tabeli w Streamlit
st.subheader("Wyekstrahowane pozycje zamówienia")
st.dataframe(df, use_container_width=True)

# 10) Przygotowanie do pobrania jako Excel
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
