# pdf_to_excel.py

import sys
import re
import io
import argparse

import pandas as pd
import PyPDF2

# --- jeśli chcesz obsługiwać trudniejsze PDF-y (jak „Wydruk.pdf”), 
# --- zainstaluj najpierw pdfplumber: pip install pdfplumber
try:
    import pdfplumber
    _HAS_PDFPLUMBER = True
except ImportError:
    _HAS_PDFPLUMBER = False


def extract_text_lines(pdf_path: str) -> list[str]:
    """
    Wyciąga wszystkie niepuste linie tekstu z PDF-a w dwóch krokach:
    1) PyPDF2
    2) (jeśli PyPDF2 nie da nic sensownego i pdfplumber jest dostępny) pdfplumber
    """
    lines = []

    # 1) Próba z PyPDF2
    try:
        with open(pdf_path, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            for page in reader.pages:
                text = page.extract_text() or ""
                for ln in text.split("\n"):
                    stripped = ln.strip()
                    if stripped:
                        lines.append(stripped)
    except Exception:
        lines = []

    # Sprawdź, czy mamy przynajmniej jedną sensowną linię (EAN lub nagłówek "Lp")
    has_ean_or_header = any(
        re.fullmatch(r"\d{13}", ln) or ln.lower().startswith("lp")
        for ln in lines
    )
    if has_ean_or_header and lines:
        return lines

    # 2) Jeśli PyPDF2 nic nie dał, spróbuj pdfplumber
    if _HAS_PDFPLUMBER:
        pl_lines = []
        try:
            with pdfplumber.open(pdf_path) as pdf:
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

    # Jeżeli dalej pusto, zwracamy to, co mamy (może to być lista pusta)
    return lines


def parse_layout_a(all_lines: list[str]) -> pd.DataFrame:
    """
    Układ A:
      - W linii: 'Kod kres.: <13-cyfrowy EAN>'
      - Następna linia: Lp (liczba)
      - Kolejne linie: fragmenty nazwy aż do wiersza z 'szt.' i ilością
    """
    produkty = []
    for idx, ln in enumerate(all_lines):
        if "Kod kres.:" in ln:
            m_ean = re.search(r"(\d{13})", ln)
            if not m_ean:
                continue
            barcode = m_ean.group(1)

            # następny wiersz to Lp
            if idx + 1 < len(all_lines) and all_lines[idx + 1].isdigit():
                lp = int(all_lines[idx + 1])
            else:
                continue

            name_parts = []
            qty = None
            j = idx + 2
            while j < len(all_lines):
                if "szt" in all_lines[j].lower():
                    # przykładowo "96,00 szt."
                    m_qty = re.search(r"([\d\s,]+)\s*szt", all_lines[j], re.IGNORECASE)
                    if m_qty:
                        qty_str = m_qty.group(1).replace(" ", "")
                        qty = float(qty_str.replace(",", "."))
                    break
                else:
                    name_parts.append(all_lines[j])
                j += 1

            name = " ".join(name_parts).strip()
            produkty.append({
                "Lp": lp,
                "Name": name,
                "Quantity": qty if qty is not None else 0,
                "Barcode": barcode
            })

    return pd.DataFrame(produkty)


def parse_layout_b(all_lines: list[str]) -> pd.DataFrame:
    """
    Układ B:
      Każda pozycja w jednej linii, np.:
        1 5029040012366 Nazwa Produktu 96,00 szt.
    Regex: ^(\d+)\s+(\d{13})\s+(.+?)\s+([\d,]+)\s+szt
    """
    produkty = []
    wzorzec = re.compile(r"^(\d+)\s+(\d{13})\s+(.+?)\s+([\d,]+)\s+szt", re.IGNORECASE)
    for ln in all_lines:
        m = wzorzec.match(ln)
        if not m:
            continue
        lp = int(m.group(1))
        barcode = m.group(2)
        name = m.group(3).strip()
        qty = float(m.group(4).replace(",", "."))
        produkty.append({
            "Lp": lp,
            "Name": name,
            "Quantity": qty,
            "Barcode": barcode
        })

    return pd.DataFrame(produkty)


def parse_layout_c(all_lines: list[str]) -> pd.DataFrame:
    """
    Układ C:
      - Linia zawiera sam 13-cyfrowy EAN
      - Następna linia: Lp
      - Kolejne wiersze: fragmenty nazwy aż do wiersza z 'szt.' i ilością
    """
    produkty = []
    idx = 0
    while idx < len(all_lines):
        ln = all_lines[idx]
        if re.fullmatch(r"\d{13}", ln):
            barcode = ln
            if idx + 1 < len(all_lines) and all_lines[idx + 1].isdigit():
                lp = int(all_lines[idx + 1])
            else:
                idx += 1
                continue

            name_parts = []
            qty = None
            j = idx + 2
            while j < len(all_lines):
                if "szt" in all_lines[j].lower():
                    m_qty = re.search(r"([\d\s,]+)\s*szt", all_lines[j], re.IGNORECASE)
                    if m_qty:
                        qty_str = m_qty.group(1).replace(" ", "")
                        qty = float(qty_str.replace(",", "."))
                    break
                else:
                    name_parts.append(all_lines[j])
                j += 1

            name = " ".join(name_parts).strip()
            produkty.append({
                "Lp": lp,
                "Name": name,
                "Quantity": qty if qty is not None else 0,
                "Barcode": barcode
            })

            idx = j + 1
        else:
            idx += 1

    return pd.DataFrame(produkty)


def detect_layout(all_lines: list[str]) -> str:
    """
    Wykrywa układ A, B lub C na podstawie:
      - B: jeżeli którakolwiek linia pasuje do wzorca B (^\d+\s+\d{13}\s+…\s+[\d,]+\s+szt)
      - C: jeżeli przynajmniej jedna linia to dokładnie 13 cyfr
      - Inaczej: A
    """
    pattern_b = re.compile(r"^\d+\s+\d{13}\s+.+\s+[\d,]+\s+szt", re.IGNORECASE)
    if any(pattern_b.match(ln) for ln in all_lines):
        return "B"
    if any(re.fullmatch(r"\d{13}", ln) for ln in all_lines):
        return "C"
    return "A"


def convert_dataframe_to_excel(df: pd.DataFrame, output_path: str) -> None:
    """
    Zapisuje cały DataFrame do jednego arkusza Excela (sheet_name="Zamowienie").
    """
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Zamowienie")


def main():
    parser = argparse.ArgumentParser(
        description="Konwerter PDF → Excel (jeden arkusz)."
    )
    parser.add_argument(
        "pdf_file",
        help="Ścieżka do wejściowego pliku PDF."
    )
    parser.add_argument(
        "excel_file",
        help="Ścieżka, gdzie zapisać wynikowy plik Excel (.xlsx)."
    )
    args = parser.parse_args()

    pdf_path = args.pdf_file
    excel_path = args.excel_file

    # 1) Wydobycie wyczyszczonych linii tekstu
    all_lines = extract_text_lines(pdf_path)
    if not all_lines:
        print(
            "Błąd: nie udało się wyciągnąć żadnych linii tekstu z PDF-a.\n"
            "Jeżeli plik jest zeskanowanym obrazem lub PyPDF2 nic nie znalazł, "
            "upewnij się, że zainstalowałeś pdfplumber (`pip install pdfplumber`), "
            "aby obsłużyć zaszyfrowane fonty/skany."
        )
        sys.exit(1)

    # 2) Wykrycie układu
    layout = detect_layout(all_lines)
    if layout == "A":
        df = parse_layout_a(all_lines)
    elif layout == "B":
        df = parse_layout_b(all_lines)
    else:  # "C"
        df = parse_layout_c(all_lines)

    # 3) Sprawdź, czy coś się znalazło
    if df.empty:
        print(
            f"Wykryto układ {layout}, ale nie znaleziono żadnych pozycji w pliku.\n"
            "Upewnij się, że plik PDF jest zgodny z którymś z obsługiwanych formatów:\n"
            "- Układ A: linia z 'Kod kres.: <EAN>' + Lp + nazwa + 'szt.'\n"
            "- Układ B: pełna pozycja w formacie '1 5029040012366 Nazwa 96,00 szt.'\n"
            "- Układ C: osobny wiersz z 13-cyfrowym EAN, potem Lp, potem nazwa, potem 'szt.'\n"
        )
        sys.exit(1)

    # 4) Zapis do Excela w jednym arkuszu
    try:
        convert_dataframe_to_excel(df, excel_path)
        print(f"Gotowe! Zapisano {len(df)} pozycji do pliku Excel: {excel_path}")
    except Exception as e:
        print(f"Błąd przy zapisie do Excela: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
