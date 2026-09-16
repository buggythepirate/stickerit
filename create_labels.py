#!/usr/bin/env python3
"""Erstellt PDF-Etiketten (2x4 pro A4-Seite) aus einer Excel-Datei.

Erwartete Spalten (Header in der ersten Zeile, Gross-/Kleinschreibung egal):
Klasse, Vorname, Nachname, Benutzername, InitialPasswort

Aufruf:
    python3 create_labels.py schueler.xlsx etiketten.pdf
"""

import sys

from openpyxl import load_workbook
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.colors import black

# Passwortrichtlinie
SONDERZEICHEN = "!@#$%^&*()_+-=[]{}|;:,.<>?/~`"
RICHTLINIE_ZEILE1 = "Passwortrichtlinie: 12 Zeichen, Gross- und Kleinschreibung,"
RICHTLINIE_ZEILE2 = f"Sonderzeichen zwingend: {SONDERZEICHEN}"

# Seitengestaltung A4
PAGE_W, PAGE_H = A4
MARGIN_X = 10 * mm
MARGIN_Y = 10 * mm
COLS = 2
ROWS = 4
GAP = 4 * mm

LABEL_W = (PAGE_W - 2 * MARGIN_X - (COLS - 1) * GAP) / COLS
LABEL_H = (PAGE_H - 2 * MARGIN_Y - (ROWS - 1) * GAP) / ROWS

PAD = 4 * mm  # Innenabstand im Etikett

# Schriftgroessen
FONTSIZE_KLASSE = 12
FONTSIZE_BODY = 10
FONTSIZE_RICHTLINIE = 7
FONTSIZE_KLEIN = 8

PW_PLACEHOLDER_COUNT = 13


def normalize(value):
    if value is None:
        return ""
    return str(value).strip()


def read_rows(xlsx_path):
    wb = load_workbook(xlsx_path, data_only=True, read_only=True)
    ws = wb.active

    rows = ws.iter_rows(values_only=True)
    try:
        header = next(rows)
    except StopIteration:
        return []

    header_map = {}
    for idx, name in enumerate(header):
        if name is None:
            continue
        header_map[str(name).strip().lower()] = idx

    needed = {
        "klasse": "Klasse",
        "vorname": "Vorname",
        "nachname": "Nachname",
        "benutzername": "Benutzername",
        "initialpasswort": "InitialPasswort",
    }
    col_idx = {}
    for key, display in needed.items():
        if key not in header_map:
            raise ValueError(f"Spalte '{display}' nicht in der Excel-Datei gefunden.")
        col_idx[key] = header_map[key]

    records = []
    for raw in rows:
        if raw is None or all(c is None or str(c).strip() == "" for c in raw):
            continue
        klasse = normalize(raw[col_idx["klasse"]])
        vorname = normalize(raw[col_idx["vorname"]])
        nachname = normalize(raw[col_idx["nachname"]])
        benutzername = normalize(raw[col_idx["benutzername"]])
        initialpasswort = normalize(raw[col_idx["initialpasswort"]])
        records.append(
            {
                "klasse": klasse,
                "vorname": vorname,
                "nachname": nachname,
                "benutzername": benutzername,
                "initialpasswort": initialpasswort,
            }
        )
    return records


def draw_label(c, x, y, rec):
    # Rahmen
    c.setStrokeColor(black)
    c.setLineWidth(0.7)
    c.rect(x, y, LABEL_W, LABEL_H)

    inner_x = x + PAD
    inner_top = y + LABEL_H - PAD

    # Klasse (oben)
    c.setFont("Helvetica-Bold", FONTSIZE_KLASSE)
    c.drawString(inner_x, inner_top - FONTSIZE_KLASSE, f"Klasse: {rec['klasse']}")

    line_h = FONTSIZE_BODY + 4
    cur = inner_top - FONTSIZE_KLASSE - line_h

    # Voller Name
    c.setFont("Helvetica-Bold", FONTSIZE_BODY)
    c.drawString(inner_x, cur, "Name:")
    c.setFont("Helvetica", FONTSIZE_BODY)
    c.drawString(inner_x + 60, cur, f"{rec['vorname']} {rec['nachname']}")
    cur -= line_h

    # Benutzername
    c.setFont("Helvetica-Bold", FONTSIZE_BODY)
    c.drawString(inner_x, cur, "Benutzername:")
    c.setFont("Helvetica", FONTSIZE_BODY)
    c.drawString(inner_x + 60, cur, rec["benutzername"])
    cur -= line_h

    # Initialpasswort
    c.setFont("Helvetica-Bold", FONTSIZE_BODY)
    c.drawString(inner_x, cur, "Initialpasswort:")
    c.setFont("Helvetica", FONTSIZE_BODY)
    c.drawString(inner_x + 60, cur, rec["initialpasswort"])
    cur -= line_h + 2

    # Vordruck neues Passwort
    c.setFont("Helvetica-Bold", FONTSIZE_KLEIN)
    c.drawString(inner_x, cur, "Neues Passwort (13 Zeichen):")
    cur -= 3 * mm

    box = 4.2 * mm
    box_gap = 1.0 * mm
    box_y = cur - box
    bx = inner_x
    for i in range(PW_PLACEHOLDER_COUNT):
        c.setStrokeColor(black)
        c.setLineWidth(0.5)
        c.rect(bx, box_y, box, box)
        bx += box + box_gap
    cur = box_y - 3 * mm

    # Passwortrichtlinie
    c.setFont("Helvetica", FONTSIZE_RICHTLINIE)
    c.drawString(inner_x, cur, RICHTLINIE_ZEILE1)
    cur -= FONTSIZE_RICHTLINIE + 1
    c.drawString(inner_x, cur, RICHTLINIE_ZEILE2)


def build_pdf(records, pdf_path):
    c = canvas.Canvas(pdf_path, pagesize=A4)

    per_page = COLS * ROWS
    for i, rec in enumerate(records):
        slot = i % per_page
        col = slot % COLS
        row = slot // COLS

        x = MARGIN_X + col * (LABEL_W + GAP)
        # von oben nach unten
        y = PAGE_H - MARGIN_Y - LABEL_H - row * (LABEL_H + GAP)

        draw_label(c, x, y, rec)

        if slot == per_page - 1 or i == len(records) - 1:
            c.showPage()

    c.save()


def main():
    if len(sys.argv) != 3:
        print("Aufruf: python3 create_labels.py <input.xlsx> <output.pdf>")
        sys.exit(1)

    xlsx_path = sys.argv[1]
    pdf_path = sys.argv[2]

    records = read_rows(xlsx_path)
    if not records:
        print("Keine Daten in der Excel-Datei gefunden.")
        sys.exit(1)

    build_pdf(records, pdf_path)
    print(f"{len(records)} Etiketten erstellt -> {pdf_path}")


if __name__ == "__main__":
    main()
