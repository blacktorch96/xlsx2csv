# -*- encoding: utf-8 -*-
import openpyxl
import csv
import tkinter as tk
from tkinter import filedialog
import os

# enc = {'Ã¼': 'ü', 'Ã¶': 'ö', 'Ã¤': 'ä', 'Ã©': 'é', 'ÃŸ': 'ß', 'Ãœ': 'Ü', 'Ã–': 'Ö', 'Ã„': 'Ä', 'Ã?': 'ß'}
enc = {'â‚¬': '€', 'â€š': '‚', 'â€ž': '„', 'â€¦': '…', 'â€¡': '‡', 'â€°': '‰', 'â€¹': '‹', 'â€˜': '‘',
       'â€™': '’', 'â€œ': '“', 'â€¢': '•', 'â€“': '–', 'â€”': '—', 'â„¢': '™', 'â€º': '›', 'Ã€': 'À',
       'Ã‚': 'Â', 'Æ’': 'ƒ', 'Ãƒ': 'Ã', 'Ã„': 'Ä', 'Ã…': 'Å', 'â€': '†', 'Ã†': 'Æ', 'Ã‡': 'Ç', 'Ë†': 'ˆ',
       'Ãˆ': 'È', 'Ã‰': 'É', 'ÃŠ': 'Ê', 'Ã‹': 'Ë', 'Å’': 'Œ', 'ÃŒ': 'Ì', 'Å½': 'Ž', 'ÃŽ': 'Î', 'Ã‘': 'Ñ',
       'Ã’': 'Ò', 'Ã“': 'Ó', 'â€': '”', 'Ã”': 'Ô', 'Ã•': 'Õ', 'Ã–': 'Ö', 'Ã—': '×', 'Ëœ': '˜', 'Ã˜': 'Ø',
       'Ã™': 'Ù', 'Å¡': 'š', 'Ãš': 'Ú', 'Ã›': 'Û', 'Å“': 'œ', 'Ãœ': 'Ü', 'Å¾': 'ž', 'Ãž': 'Þ', 'Å¸': 'Ÿ',
       'ÃŸ': 'ß', 'Â¡': '¡', 'Ã¡': 'á', 'Â¢': '¢', 'Ã¢': 'â', 'Â£': '£', 'Ã£': 'ã', 'Â¤': '¤', 'Ã¤': 'ä',
       'Â¥': '¥', 'Ã¥': 'å', 'Â¦': '¦', 'Ã¦': 'æ', 'Â§': '§', 'Ã§': 'ç', 'Â¨': '¨', 'Ã¨': 'è', 'Â©': '©',
       'Ã©': 'é', 'Âª': 'ª', 'Ãª': 'ê', 'Â«': '«', 'Ã«': 'ë', 'Â¬': '¬', 'Ã¬': 'ì', 'Â­': '­', 'Ã­': 'í',
       'Â®': '®', 'Ã®': 'î', 'Â¯': '¯', 'Ã¯': 'ï', 'Â°': '°', 'Ã°': 'ð', 'Â±': '±', 'Ã±': 'ñ', 'Â²': '²',
       'Ã²': 'ò', 'Â³': '³', 'Ã³': 'ó', 'Â´': '´', 'Ã´': 'ô', 'Âµ': 'µ', 'Ãµ': 'õ', 'Â¶': '¶', 'Ã¶': 'ö',
       'Â·': '·', 'Ã·': '÷', 'Â¸': '¸', 'Ã¸': 'ø', 'Â¹': '¹', 'Ã¹': 'ù', 'Âº': 'º', 'Ãº': 'ú', 'Â»': '»',
       'Ã»': 'û', 'Â¼': '¼', 'Ã¼': 'ü', 'Â½': '½', 'Ã½': 'ý', 'Â¾': '¾', 'Ã¾': 'þ', 'Â¿': '¿', 'Ã¿': 'ÿ',
       'Ã': 'Á', 'Å': 'Š', 'Ã': 'Í', 'Ã': 'Ï', 'Ã': 'Ð', 'Ã': 'Ý', 'Â': '', 'Ã': 'à'}
# https://www.i18nqa.com/debug/utf8-debug.html
encoding_issue = 0
encoding_solved = 0


def celling(xcell) -> str:
    value = str(xcell.value).replace('\n', ' ').replace('\t', ' ').replace('\r', ' ').replace("None", '')
    if chr(0xC3) in value or chr(0xC2) in value or chr(0xC5) in value or chr(0xE2) in value or chr(0xCB) in value or chr(0xC6) in value:
        global encoding_issue, encoding_solved
        for k, v in enc.items():
            if k in value:
                value = value.replace(k, v)
                encoding_solved += 1
        if 'Ã' in value:
            encoding_issue += 1
            # todo: sind noch weitere Endcodierungsfehler vorhanden?
            print("weiterer Encodierungsfehler!")
    if xcell.data_type == xcell.TYPE_NUMERIC and xcell.number_format == '"0"0':  # Sonderfall #1 Telefonnummern
        value = "0" + value
    return value


root = tk.Tk()
root.withdraw()

load_file = filedialog.askopenfilename(title="Excel-Datei zum konvertieren auswählen",
                                       filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))
if load_file == "":
    print("Keine Datei angegeben. Abbruch.")
    exit()

file_root, file_ext = os.path.splitext(load_file)
save_file = file_root + ".csv"
# save_name = str(load_file).replace(str(load_file).rsplit('.', 2)[1], 'csv')
# save_file = filedialog.asksaveasfilename(title="Zieldatei auswählen (CSV)", initialfile=save_name,
#                                         filetypes=(("csv files", "*.csv"), ("all files", "*.*")))
# if save_file == "":
#    print("Keine Datei zum Speichern angegeben. Abbruch.")
#    exit()

print("Öffne XLSX: {}\nDies kann je nach Dateigröße ein paar Minuten dauern...".format(load_file))
wb = openpyxl.load_workbook(load_file)
# sh = wb.get_active_sheet()
sh = wb.active

print("Schreibe CSV: {}".format(save_file))
with open(save_file, 'w', newline="", encoding="utf-8") as f:
    c = csv.writer(f, delimiter='\t', quoting=csv.QUOTE_ALL)
    i = 0
    m = sh.max_row
    p = 0

    for r in sh.rows:
        p2 = int((i / m) * 100)
        i += 1
        p = int((i / m) * 100)
        if p != p2 and p % 10 == 0:
            print("{} %".format(p).rjust(5, ' '))
        # print(r)
        c.writerow(
            [celling(cell) for cell in r])

wb.close()

if encoding_solved > 0:
    print(f"Es wurden {encoding_solved} Encodierungsfehler behoben. Bitte beachten! ")
if encoding_issue > 0:
    print(f"Es wurden {encoding_issue} Encodierungsfehler nicht gelöst! Bitte beachten! ")

print("Die Datei kann nun in Salesforce importiert werden:")
print("z.B. DataImporter: Import-Typ: 'CSV', Zeichencode 'UTF-8' und Werte getrennt: 'Registerkarte'")
print("Achtung! Das Nachbearbeiten mit Excel kann die Datei zerstören!")
print("done")
