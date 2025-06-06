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
       - **Układ D**: proste linie zawierające tylko EAN (13 cyfr) i ilość, np.  
         `5029040012366 Nazwa Produktu 96,00 szt.` lub `5029040012366 96,00 szt.`  
       - **Układ E**: linie zaczynające się od Lp i nazwy, potem ilość, a poniżej „Kod kres.: <EAN>”.  
         (Przykłady plików typu `Gussto wola park.pdf` czy `Zamówienie nr ZD 0175_05_25.pdf`.)  
       - **Układ B**: każda pozycja w jednej linii, np.  
         `<Lp> <EAN(13)> <pełna nazwa> <ilość>,<xx> szt.`  
       - **Układ C**: czysty 13-cyfrowy EAN w osobnej linii, potem Lp w osobnej linii, potem nazwa, „szt.” i ilość.  
       - **Układ A**: „Kod kres.: <EAN>” w osobnej linii, Lp w osobnej linii, fragmenty nazwy przed i po kolumnie cen/ilości.
    4. Wywołuje odpowiedni parser i wyświetla wynik w formie tabeli (`Lp`, `Name`, `Quantity`, `Barcode`).
    5. Umożliwia pobranie danych jako plik Excel.
    """
)


# ──────────────────────────────────────────────────────────────────────────────
# 1) POMOCNICZE FUNKCJE DO WYCIĄGANIA TEKSTU

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


# ──────────────────────────────────────────────────────────────────────────────
# 2) PARSERY UKŁADÓW „NOWYCH” (D i E) ORAZ „STANDARDOWYCH” (B, C, A)
#    (te funkcje działają na liście linii, bez względu na to,
#     czy linie pochodzą z PyPDF2, czy z pdfplumber)

def parse_layout_d(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla układu D – proste linie zawierające EAN (13 cyfr) i ilość w formacie "<ilość>,<xx> szt.".
    Przykład:
      5029040012366 Nazwa Produktu 96,00 szt.
      5029040012403 96,00 szt.
    - Lp automatycznie rośnie od 1.
    - Name pozostaje puste.
    """
    products = []
    pattern = re.compile(
        r"^(\d{13})(?:\s+.*?)*\s+(\d{1,3}),\d{2}\s+szt",
        flags=re.IGNORECASE
    )
    lp_counter = 1
    for ln in all_lines:
        m = pattern.match(ln)
        if m:
            barcode_val = m.group(1)
            qty_val = int(m.group(2).replace(" ", "").replace(" ", ""))
            products.append({
                "Lp": lp_counter,
                "Name": "",
                "Quantity": qty_val,
                "Barcode": barcode_val
            })
            lp_counter += 1
    return pd.DataFrame(products)


def parse_layout_e(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla układu E – linia z Lp i tekstem nazwy oraz ilością w tej samej linii,
    a poniżej (ewentualnie po liniach typu "ARA...") znajduje się linia "Kod kres.: <EAN>".
    
    Przykładowa sekwencja:
      1 CANAGAN Kot 0,375kg 8 szt. …
      ARA000585
      Kod kres.: 5029040013097
      2 CANAGAN Kot SCOTTISH 8 szt. …
      Run Turkey
      ARA000613
      Kod kres.: 5029040013318
      ...
    
    Logika:
      - Wzorzec dopasowujący Lp na początku, potem dowolny tekst (nazwa),
        a za nim „<ilość> szt.” w tej samej linii.
      - Po linii z Lp zbieramy kolejne wiersze aż do znalezienia „Kod kres.:”:
        • Jeśli linia jest czysto alfanumeryczna bez spacji (np. "ARA000613"), pomijamy.
        • W przeciwnym razie traktujemy kolejną linię jako część nazwy.
      - Gdy napotkamy „Kod kres.: <EAN>”, wyciągamy EAN i kończymy tę pozycję.
    """
    products = []
    i = 0
    pattern_item = re.compile(r"^(\d+)\s+(.+?)\s+(\d{1,3})\s+szt\.", flags=re.IGNORECASE)

    while i < len(all_lines):
        ln = all_lines[i]
        m = pattern_item.match(ln)
        if m:
            lp_val = int(m.group(1))
            initial_name = m.group(2).strip()
            qty_val = int(m.group(3))
            name_parts = [initial_name]
            barcode_val = None

            j = i + 1
            while j < len(all_lines):
                next_ln = all_lines[j]

                if next_ln.lower().startswith("kod kres"):
                    parts = next_ln.split(":", 1)
                    if len(parts) == 2:
                        barcode_val = parts[1].strip()
                    j += 1
                    break

                if re.fullmatch(r"[A-Za-z0-9]+", next_ln):
                    # linia katalogu (ARA...), pomijamy
                    j += 1
                    continue

                # w przeciwnym razie traktujemy jako fragment nazwy
                name_parts.append(next_ln.strip())
                j += 1

            full_name = " ".join(name_parts).strip()
            products.append({
                "Lp": lp_val,
                "Name": full_name,
                "Quantity": qty_val,
                "Barcode": barcode_val
            })

            i = j
        else:
            i += 1

    return pd.DataFrame(products)


def parse_layout_b(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla układu B – każda pozycja w jednej linii:
      <Lp> <EAN(13)> <pełna nazwa> <ilość>,<xx> szt. …
    Wyciąga Lp, Barcode, Name, Quantity.
    Przykład:
      3 5029040012045 Canalban Kot … 12,00 szt.
    """
    products = []
    pattern = re.compile(
        r"^(\d+)\s+(\d{13})\s+(.+?)\s+(\d{1,3}),\d{2}\s+szt",
        flags=re.IGNORECASE
    )
    for ln in all_lines:
        m = pattern.match(ln)
        if m:
            lp_val = int(m.group(1))
            barcode_val = m.group(2)
            name_val = m.group(3).strip()
            qty_val = int(m.group(4).replace(" ", "").replace(" ", ""))
            products.append({
                "Lp": lp_val,
                "Name": name_val,
                "Quantity": qty_val,
                "Barcode": barcode_val
            })
    return pd.DataFrame(products)


def parse_layout_c(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla układu C – czysty 13-cyfrowy EAN w osobnej linii, potem Lp, potem nazwa,
    potem "szt." i ilość w kolejnych wierszach.
    
    Przykład:
      5029040012366
      3
      Nazwa Produktu
      szt.
      12
      (opcjonalnie: "Kod kres.: <...>")
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
        barcode_val = all_lines[max(eans_before)] if eans_before else None

        name_val = all_lines[lp_idx + 1] if lp_idx + 1 < len(all_lines) else None

        qty_val = None
        for j in range(lp_idx + 1, len(all_lines) - 2):
            if all_lines[j].lower() == "szt." and re.fullmatch(r"\d+", all_lines[j + 2]):
                qty_val = int(all_lines[j + 2])
                break

        if name_val and qty_val is not None:
            products.append({
                "Lp": int(all_lines[lp_idx]),
                "Name": name_val.strip(),
                "Quantity": qty_val,
                "Barcode": barcode_val
            })

    return pd.DataFrame(products)


def parse_layout_a(all_lines: list[str]) -> pd.DataFrame:
    """
    Parser dla układu A – "Kod kres.: <EAN>" w osobnej linii, Lp w osobnej linii,
    fragmenty nazwy przed i po kolumnie cen/ilości.

    Przykład fragmentu:
      1
      Nazwa Produktu
      …
      8
      szt.
      Kod kres.: 5029040013097
      (kolejna pozycja)
    """
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

    idx_ean = [i for i, ln in enumerate(all_lines) if ln.lower().startswith("kod kres")]
    products = []
    for idx, lp_idx in enumerate(idx_lp):
        prev_lp = idx_lp[idx - 1] if idx > 0 else -1
        next_lp = idx_lp[idx + 1] if idx + 1 < len(idx_lp) else len(all_lines)

        valid_eans = [e for e in idx_ean if prev_lp < e < next_lp]
        barcode_val = None
        if valid_eans:
            parts = all_lines[max(valid_eans)].split(":", 1)
            if len(parts) == 2:
                barcode_val = parts[1].strip()

        name_parts: list[str] = []
        qty_val = None
        qty_idx = None

        for j in range(lp_idx + 1, next_lp):
            ln = all_lines[j]
            if re.fullmatch(r"\d+", ln) and (j + 1 < next_lp and all_lines[j + 1].lower() == "szt."):
                qty_idx = j
                qty_val = int(ln)
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

        full_name = " ".join(name_parts).strip()
        products.append({
            "Lp": int(all_lines[lp_idx]),
            "Name": full_name,
            "Quantity": qty_val,
            "Barcode": barcode_val
        })

    return pd.DataFrame(products)


# ──────────────────────────────────────────────────────────────────────────────
# 3) GŁÓWNA LOGIKA: WCZYTANIE PLIKÓW I WYBOR PARSERA

# 3.1) Wgraj PDF
uploaded_file = st.file_uploader("Wybierz plik PDF ze zamówieniem", type=["pdf"])
if uploaded_file is None:
    st.info("Proszę wgrać plik PDF, aby kontynuować.")
    st.stop()

pdf_bytes = uploaded_file.read()

# 3.2) Próba wydobycia tekstu przez PyPDF2 (stara metoda)
lines_py = extract_text_with_pypdf2(pdf_bytes)

# 3.3) Rozpoznaj układ D i E w tekście z PyPDF2,
#      aby wiedzieć, czy należy użyć „nowych” parserów.
pattern_d = re.compile(r"^\d{13}(?:\s+.*?)*\s+\d{1,3},\d{2}\s+szt", flags=re.IGNORECASE)
is_layout_d_py = any(pattern_d.match(ln) for ln in lines_py)

pattern_e = re.compile(r"^\d+\s+.+?\s+\d{1,3}\s+szt\.", flags=re.IGNORECASE)
has_kod_kres_py = any(ln.lower().startswith("kod kres") for ln in lines_py)
is_layout_e_py = any(pattern_e.match(ln) for ln in lines_py) and has_kod_kres_py

# 3.4) Jeśli PyPDF2 wyciągnęło jakieś linie i nie są to układy D/E,
#      to użyjemy „starego” kodu parsującego (układy B, C lub A).
df = pd.DataFrame()
if lines_py and not is_layout_d_py and not is_layout_e_py:
    # Spróbuj wykryć układ B (Lp + EAN + nazwa + ilość w jednej linii)
    pattern_b = re.compile(r"^\d+\s+\d{13}\s+.+\s+\d{1,3},\d{2}\s+szt", flags=re.IGNORECASE)
    is_layout_b_py = any(pattern_b.match(ln) for ln in lines_py)

    # Spróbuj wykryć układ C (czysty 13-cyfrowy EAN w oddzielnej linii, ale nie układ B)
    has_pure_ean_py = any(re.fullmatch(r"\d{13}", ln) for ln in lines_py)
    is_layout_c_py = has_pure_ean_py and not is_layout_b_py

    if is_layout_b_py:
        df = parse_layout_b(lines_py)
    elif is_layout_c_py:
        df = parse_layout_c(lines_py)
    else:
        # Domyślny parser A, tak jak w oryginalnym kodzie
        df = parse_layout_a(lines_py)

# 3.5) Jeśli nic nie znaleziono starym sposobem lub linie PyPDF2 wskazują na układ D/E,
#      to przejdźmy do „nowego” wydobywania tekstu przez pdfplumber
if df.empty:
    lines_new = extract_text_with_pdfplumber(pdf_bytes)

    # Upewnijmy się, że w ogóle jest tekst
    if not lines_new:
        st.error(
            "Nie udało się wyciągnąć tekstu z tego PDF-a. "
            "Być może wymaga OCR – wykonaj OCR (np. Tesseract) i wgraj ponownie."
        )
        st.stop()

    # Ponownie rozpoznaj układ D/E, ale już dla pdfplumber
    is_layout_d_new = any(pattern_d.match(ln) for ln in lines_new)
    has_kod_kres_new = any(ln.lower().startswith("kod kres") for ln in lines_new)
    is_layout_e_new = any(pattern_e.match(ln) for ln in lines_new) and has_kod_kres_new

    # Rozpoznaj układ B i C w liniach z pdfplumber (jeśli potrzebne)
    is_layout_b_new = any(re.compile(r"^\d+\s+\d{13}\s+.+\s+\d{1,3},\d{2}\s+szt", flags=re.IGNORECASE).match(ln) for ln in lines_new)
    has_pure_ean_new = any(re.fullmatch(r"\d{13}", ln) for ln in lines_new)
    is_layout_c_new = has_pure_ean_new and not is_layout_b_new

    # Wybór parsera dla „nowego” trybu
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

# 3.6) Odfiltruj wiersze, które nie mają wartości Quantity (jeśli kolumna istnieje)
if "Quantity" in df.columns:
    df = df.dropna(subset=["Quantity"]).reset_index(drop=True)

# 3.7) Sprawdź, czy cokolwiek zostało wyciągnięte
if df.empty:
    st.error(
        "Po parsowaniu nie znaleziono pozycji zamówienia. "
        "Upewnij się, że PDF zawiera kody EAN oraz ilości w formacie rozpoznawalnym przez parser."
    )
    st.stop()

# ──────────────────────────────────────────────────────────────────────────────
# 4) WYŚWIETLENIE WYNIKÓW I POBRANIE EXCEL

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
