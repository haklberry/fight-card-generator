#!/usr/bin/env python3
"""Creates sample_fights.xlsx and template_zapasy.xlsx with the new structure.
Event info (title/date/ring) is stored once in the header section, NOT repeated per fight row.
"""

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
except ImportError:
    import sys
    sys.exit("Run: pip install openpyxl")

TITLE = "149. KOLO NÁRODNÍ LIGY"
DATE  = "18/10/2025"
RING  = "RING A"

NAVY  = "1B2A4A"
RED   = "E8374A"
WHITE = "FFFFFF"
LGREY = "D0D4DC"
MGREY = "8A8F9E"

FIGHT_HEADERS = [
    "fight_number", "red_name", "red_gym", "red_weight",
    "time", "blue_weight", "blue_name", "blue_gym", "note", "section",
]

COL_WIDTHS = [8, 26, 30, 8, 10, 8, 26, 32, 22, 20]   # A–J

FIGHTS = [
    # (fight_number, red_name, red_gym, red_weight, time, blue_weight, blue_name, blue_gym, note, section)
    (1,  "Michael Berka",      "Hanuman Gym Praha",          75, "3x2 min", 75, "Tadeáš Míka",          "Gladiators Gym Třeboň",          "",               "Předzápasy"),
    (2,  "Patrik Starovic",    "Alpha Gym Karlovy Vary",     67, "3x2 min", 67, "Filip Gabriel",         "Spejbl Gym Praha",               "",               ""),
    (3,  "Jakub Marvan",       "Adamas Team",                75, "3x2 min", 75, "Richard Novotný",       "Alpha Gym Karlovy Vary",          "",               ""),
    (4,  "Ion Moraru",         "Titan Boxing Gym Z.S.",      81, "3x2 min", 81, "Shoham Niv",            "Spejbl Gym Praha",               "",               ""),
    (5,  "Radim Oranský",      "Gladiators Gym Č. B.",       86, "3x2 min", 86, "Adam Voborný",          "Jetsaam Gym",                    "",               ""),
    (6,  "Švehlová Viktorie",  "Ravána Gym, Z.S.",           57, "3x2 min", 60, "Hájková Nikoleta",      "SK Pretorian",                   "děti, chrániče", "Hlavní karta"),
    (7,  "Natalia Kopyn",      "Hanuman Gym Praha",          48, "3x2 min", 54, "Thea Honegrová",        "BT Gym Praha",                   "",               ""),
    (8,  "Jan Maják",          "Better Club Z.S.",           60, "3x2 min", 60, "Pernička David",        "BT Gym Praha",                   "",               ""),
    (9,  "Filip Hrabačka",     "Yaksha Gym Plzeň",           63, "3x2 min", 63, "Jaroslav Butkaj",       "BT Gym Praha",                   "",               ""),
    (10, "Jakub Černý",        "Spejbl Gym Praha",           67, "3x2 min", 67, "Tomáš Bernas",          "Muaythai Chomutov Zak's Team",   "",               ""),
    (11, "Matěj Chylík",       "Lanna Gym Praha",            67, "3x2 min", 67, "Lukáš Trifanov",        "Fight Club Peňáz Z.S.",          "",               ""),
    (12, "Maksym Khodakivsky", "Adamas Team",                67, "3x2 min", 67, "Richard Bazal",         "Aplik Muay Thai -AMT, Z.S.",     "",               ""),
    (13, "Tadeáš Odipe",       "Lanna Gym Praha",            71, "3x2 min", 71, "Pavel Flaška",          "Hanuman Gym Praha",               "",               ""),
    (14, "Šimon Hadrava",      "Wu-Shu Pelhřimov",           71, "3x2 min", 71, "Michael Ševčík",        "Rasgym",                         "",               ""),
    (15, "Tymur Kolesnik",     "Titan Boxing Gym Z.S.",      81, "3x2 min", 81, "Alexandr Fromelius",    "Lanna Gym Praha",                "",               ""),
]


def navy_fill():
    return PatternFill("solid", fgColor=NAVY)

def red_fill():
    return PatternFill("solid", fgColor=RED)

def grey_fill():
    return PatternFill("solid", fgColor=LGREY)

def white_fill():
    return PatternFill("solid", fgColor=WHITE)


def build_sheet(ws, fights_data, is_template=False):
    """Write the event info section + fight table onto ws."""

    # ── Row 1: Event section header ──────────────────────────────────────────
    ws.append(["INFORMACE O AKCI — vyplň jednou"] + [""] * 8)
    ws.merge_cells("A1:I1")
    c = ws["A1"]
    c.font      = Font(bold=True, color=WHITE, size=13)
    c.fill      = red_fill()
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    # ── Row 2: Event labels ───────────────────────────────────────────────────
    labels = [
        "Název akce [title]",
        "Datum [date]",
        "Ring [ring]",
    ]
    ws.append(labels + [""] * 6)
    for col_i, lbl in enumerate(labels, start=1):
        c = ws.cell(row=2, column=col_i)
        c.font      = Font(bold=True, color=WHITE, size=10)
        c.fill      = navy_fill()
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 18

    # ── Row 3: Event values ───────────────────────────────────────────────────
    if is_template:
        values = [TITLE, DATE, RING]   # sample values as placeholder
    else:
        values = [TITLE, DATE, RING]
    ws.append(values + [""] * 6)
    for col_i in range(1, 4):
        c = ws.cell(row=3, column=col_i)
        c.font      = Font(bold=True, size=11)
        c.fill      = white_fill()
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[3].height = 22

    # ── Row 4: Blank separator ────────────────────────────────────────────────
    ws.append([""] * 9)
    ws.row_dimensions[4].height = 8

    # ── Row 5: Fight table header ─────────────────────────────────────────────
    ws.append(FIGHT_HEADERS)
    for col_i in range(1, len(FIGHT_HEADERS) + 1):
        c = ws.cell(row=5, column=col_i)
        c.font      = Font(bold=True, color=WHITE, size=10)
        c.fill      = navy_fill()
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[5].height = 18

    # ── Row 6: Hint row (fight_number not numeric → skipped by load_excel) ───
    hints = ["→", "Jméno Příjmení", "Název Gymu", "75", "3x2 min", "75",
             "Jméno Příjmení", "Název Gymu", "poznámka (nepovinné)", "Název sekce (nepovinné)"]
    ws.append(hints)
    for col_i in range(1, 11):
        c = ws.cell(row=6, column=col_i)
        c.font      = Font(italic=True, color=MGREY, size=9)
        c.fill      = grey_fill()
        c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[6].height = 16

    # ── Rows 7+: Fight data ───────────────────────────────────────────────────
    if is_template:
        # Pre-fill numbers 1–15 in grey; all other cells blank
        for num in range(1, 16):
            row_data = [num] + [""] * 8
            ws.append(row_data)
            data_row = 6 + num
            ws.cell(row=data_row, column=1).font = Font(color=LGREY, size=10)
            ws.row_dimensions[data_row].height = 18
    else:
        for row_tuple in fights_data:
            ws.append(list(row_tuple))
            data_row = ws.max_row
            ws.row_dimensions[data_row].height = 18

    # ── Column widths ─────────────────────────────────────────────────────────
    for col_letter, width in zip("ABCDEFGHI", COL_WIDTHS):
        ws.column_dimensions[col_letter].width = width

    # ── Legend ────────────────────────────────────────────────────────────────
    legend_start = ws.max_row + 2
    ws.cell(row=legend_start, column=1, value="LEGENDA")
    ws.cell(row=legend_start, column=1).font = Font(bold=True, color=NAVY, size=10)
    entries = [
        ("fight_number", "Číslo zápasu (1, 2, 3 …)"),
        ("red_name / blue_name", "Celé jméno závodníka"),
        ("red_gym / blue_gym", "Název gymu"),
        ("red_weight / blue_weight", "Váha v kg (bez jednotky)"),
        ("time", "Délka zápasu, např. 3x2 min"),
        ("note", "Nepovinná poznámka — zobrazí se uprostřed karty (např. děti, chrániče)"),
    ]
    for i, (col_name, desc) in enumerate(entries, start=legend_start + 1):
        ws.cell(row=i, column=1, value=col_name).font = Font(bold=True, size=9)
        ws.cell(row=i, column=2, value=desc).font     = Font(size=9)


def make_sample():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Fights"
    build_sheet(ws, FIGHTS, is_template=False)
    wb.save("sample_fights.xlsx")
    print("Created sample_fights.xlsx")


def make_template():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Fights"
    build_sheet(ws, [], is_template=True)
    wb.save("template_zapasy.xlsx")
    print("Created template_zapasy.xlsx")


if __name__ == "__main__":
    make_sample()
    make_template()
