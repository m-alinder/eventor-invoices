import re
import os
import argparse
import pdfplumber
import pandas as pd


# ---------------------------------------------------------
# Kommandoradsargument
# ---------------------------------------------------------

parser = argparse.ArgumentParser(description="Extrahera fakturadata från PDF-filer.")
parser.add_argument("-i", "--input", required=True, help="Mapp där PDF-fakturorna finns")
parser.add_argument("-o", "--output", required=True, help="Excel-fil som ska skapas")

args = parser.parse_args()

INPUT_FOLDER = args.input
OUTPUT_FILE = args.output


# ---------------------------------------------------------
# Hjälpfunktioner
# ---------------------------------------------------------

def extract_header_info(lines):
    """Hämtar tävlingsdatum och tävlingsnamn från rekonstrukterade rader."""
    date = ""
    event = ""

    for line in lines:
        m_date = re.search(r"Tävlingsdatum\s+(\d{4}-\d{2}-\d{2})", line)
        if m_date:
            date = m_date.group(1)

        m_event = re.search(r"Tävling\s+(.+)", line)
        if m_event:
            event = m_event.group(1).strip()

    return date, event


def parse_benamning(text):
    """
    Parsa benämningen enligt reglerna:
    1) <Tjänst> för <Namn> i <Klass>   (klass kan innehålla mellanslag)
    2) <Tjänst> för <Namn>
    """

    t = " ".join(text.split())  # normalisera mellanslag

    # 1) <Tjänst> för <Namn> i <Klass>
    m1 = re.match(r"(.+?)\s+för\s+(.+?)\s+i\s+(.+)$", t)
    if m1:
        tjänst = m1.group(1).strip()
        namn = m1.group(2).strip()
        klass = m1.group(3).strip()
        return tjänst, namn, klass

    # 2) <Tjänst> för <Namn>
    m2 = re.match(r"(.+?)\s+för\s+(.+)$", t)
    if m2:
        tjänst = m2.group(1).strip()
        namn = m2.group(2).strip()
        klass = ""
        return tjänst, namn, klass

    return None, None, None


def split_line_from_right(line):
    """
    Plocka ut belopp, pris, antal från höger:
    ... <benämning> <antal> <pris> <belopp>
    Både pris och belopp är 'NNN SEK'.
    Returnerar (benämning, antal, pris, belopp) eller (None, None, None, None) vid fel.
    """

    s = line.strip()

    # 1) Hitta belopp (sista NNN SEK)
    m_belopp = re.search(r"(\d+)\s*SEK\s*$", s)
    if not m_belopp:
        return None, None, None, None
    belopp_str = m_belopp.group(0)
    belopp = m_belopp.group(1)
    s = s[:m_belopp.start()].rstrip()

    # 2) Hitta pris (nästa NNN SEK från höger)
    m_pris = re.search(r"(\d+)\s*SEK\s*$", s)
    if not m_pris:
        return None, None, None, None
    pris_str = m_pris.group(0)
    pris = m_pris.group(1)
    s = s[:m_pris.start()].rstrip()

    # 3) Hitta antal (sista ensamma talet)
    m_antal = re.search(r"(\d+)\s*$", s)
    if not m_antal:
        return None, None, None, None
    antal = m_antal.group(1)
    benamning = s[:m_antal.start()].rstrip()

    return benamning, antal, f"{pris} SEK", f"{belopp} SEK"


# ---------------------------------------------------------
# Körning
# ---------------------------------------------------------

valid_rows = []
invalid_rows = []

for filename in os.listdir(INPUT_FOLDER):
    if not filename.lower().endswith(".pdf"):
        continue

    pdf_path = os.path.join(INPUT_FOLDER, filename)

    with pdfplumber.open(pdf_path) as pdf:
        # Header-info från första sidan
        first_text = pdf.pages[0].extract_text(x_tolerance=2)
        lines_page1 = first_text.split("\n") if first_text else []
        datum, tävling = extract_header_info(lines_page1)

        # Gå igenom alla sidor och alla rader
        for page in pdf.pages:
            raw_text = page.extract_text(x_tolerance=2)
            if not raw_text:
                continue

            for line in raw_text.split("\n"):
                # Vi bryr oss bara om rader med belopp (NNN SEK i slutet)
                if not re.search(r"\d+\s*SEK\s*$", line):
                    continue

                benamning, antal, pris, belopp = split_line_from_right(line)
                if benamning is None:
                    # Vi fick inte ut kolumnerna – logga som felaktig
                    invalid_rows.append({
                        "Filnamn": filename,
                        "Datum": datum,
                        "Tävling": tävling,
                        "Klass": "",
                        "Namn": "",
                        "Tjänst": line.strip(),
                        "Avgift": ""
                    })
                    continue

                # Extrahera avgift (beloppet i siffror)
                m_price_val = re.search(r"(\d+)", belopp)
                if not m_price_val:
                    continue
                avgift = int(m_price_val.group(1))

                # Parsa benämningen till tjänst/namn/klass
                tjänst, namn, klass = parse_benamning(benamning)

                if tjänst is None:
                    invalid_rows.append({
                        "Filnamn": filename,
                        "Datum": datum,
                        "Tävling": tävling,
                        "Klass": "",
                        "Namn": "",
                        "Tjänst": benamning,
                        "Avgift": avgift
                    })
                else:
                    valid_rows.append({
                        "Filnamn": filename,
                        "Datum": datum,
                        "Tävling": tävling,
                        "Klass": klass,
                        "Namn": namn,
                        "Tjänst": tjänst,
                        "Avgift": avgift
                    })


# ---------------------------------------------------------
# Skapa Excel
# ---------------------------------------------------------

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    pd.DataFrame(valid_rows).to_excel(writer, sheet_name="Data", index=False)
    pd.DataFrame(invalid_rows).to_excel(writer, sheet_name="Felaktiga rader", index=False)

print(f"Klart! Excel-fil skapad: {OUTPUT_FILE}")
