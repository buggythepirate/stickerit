# stickerit

Erzeugt PDF-Etiketten (2x4 pro A4-Seite) aus einer Excel-Datei.

## Nutzung

```bash
pip install openpyxl reportlab
python3 create_labels.py schueler.xlsx etiketten.pdf
```

Erwartete Excel-Spalten (erste Zeile als Header):
`Klasse | Vorname | Nachname | Benutzername | InitialPasswort`

Pro Excel-Zeile entsteht ein umrandetes Etikett mit Klasse, vollem Namen,
Benutzername, Initialpasswort, Vordruck für ein neues Passwort (13 Kästchen)
und der Passwortrichtlinie.
