import re
import pandas as pd
import PyPDF2

def parse_pdf_to_dataframe(reader: PyPDF2.PdfReader) -> pd.DataFrame:
    """
    Parsuje PyPDF2.PdfReader (zamówienie PDF) w taki sposób, żeby:
    - wykrywać bloki pozycji na wielu stronach,
    - odrzucać stopki ("Strona ..."),
    - a przede wszystkim: gdy "Kod kres." wpadnie nad lub pod nagłówkiem "Lp",
      przypisać go poprawnie do bieżącej pozycji (current).
    """
    products = []
    current = None         # słownik {'Lp':..., 'Name':..., 'Quantity':..., 'Barcode':...}
    capture_name = False   # czy wciąż zbieramy fragmenty nazwy
    name_lines = []        # lista pośrednia do składania nazwy

    for page in reader.pages:
        raw_lines = page.extract_text().split("\n")

        # 1) Odcinamy stopkę: każda linia od momentu, gdy pojawi się słowo "Strona"
        footer_idx = None
        for i, ln in enumerate(raw_lines):
            if "Strona" in ln:   # jeśli w stopce jest inne słowo-klucz, dopisz je tutaj
                footer_idx = i
                break

        if footer_idx is not None:
            lines = raw_lines[:footer_idx]
        else:
            lines = raw_lines

        # 2) Szukamy nagłówka "Lp" na bieżącej stronie (po odcięciu stopki)
        header_idx = None
        for i, ln in enumerate(lines):
            if ln.strip().startswith("Lp"):
                header_idx = i
                break

        # 3) Jeśli znaleźliśmy nagłówek, to najpierw przeszukujemy wszystki wiersze przed header_idx
        #    tylko w kontekście szukania "Kod kres." dla ostatniego current.
        if header_idx is not None:
            # a) pętlą po wszystkich wierszach od 0 do header_idx-1:
            for ln in lines[:header_idx]:
                stripped = ln.strip()
                if stripped.startswith("Kod kres."):
                    parts = stripped.split(":", maxsplit=1)
                    if len(parts) == 2 and current is not None:
                        barcode = parts[1].strip()
                        if current.get("Barcode") is None:
                            current["Barcode"] = barcode
            # b) Ustawiamy start_idx dopiero za headerem:
            start_idx = header_idx + 1
        else:
            # Jeżeli na tej stronie nie ma powtórzonego nagłówka "Lp", to kontynuujemy od początku:
            start_idx = 0

        # 4) Dalej: parsowanie linii od start_idx do końca "lines"
        for i in range(start_idx, len(lines)):
            stripped = lines[i].strip()

            # 4a) Jeżeli wiersz to "Kod kres.: XXXXX", przypiszemy go do current (jeśli jest puste)
            if stripped.startswith("Kod kres."):
                parts = stripped.split(":", maxsplit=1)
                if len(parts) == 2 and current is not None:
                    barcode = parts[1].strip()
                    if current.get("Barcode") is None:
                        current["Barcode"] = barcode
                continue

            # 4b) Jeżeli wiersz to wyłącznie liczba (np. "12", "150" itp.)
            if re.fullmatch(r"\d+", stripped):
                # 4b-i) Jeżeli kolejny wiersz to "szt.", to jest to ilość (Quantity)
                if i + 1 < len(lines) and lines[i + 1].strip().lower() == "szt.":
                    qty = int(stripped)
                    if current is not None:
                        current["Quantity"] = qty
                        # Sklejamy nazwę z fragmentów
                        full_name = " ".join(name_lines).strip()
                        current["Name"] = full_name
                        name_lines = []
                        capture_name = False
                    continue
                else:
                    # 4b-ii) W przeciwnym wypadku to nowy Lp → rozpoczynamy nową pozycję
                    lp_number = int(stripped)
                    current = {"Lp": lp_number, "Name": None, "Quantity": None, "Barcode": None}
                    products.append(current)
                    capture_name = True
                    name_lines = []
                    continue

            # 4c) Jeżeli capture_name=True i wiersz nie-pusty → składamy kolejne fragmenty nazwy
            if capture_name and stripped:
                name_lines.append(stripped)
                continue

            # Pozostałe wiersze (np. ceny, VAT, puste itp.) ignorujemy

        # Po zakończeniu bieżącej strony – kontynuujemy do następnej, z zachowaniem
        # aktualnego `current`, `capture_name` i `name_lines`.

    # 5) Po przejściu WSZYSTKICH stron budujemy DataFrame i filtrujemy niekompletne wiersze:
    df = pd.DataFrame(products)
    df = df.dropna(subset=["Name", "Quantity"]).reset_index(drop=True)
    return df
