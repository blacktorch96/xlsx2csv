# -*- encoding: utf-8 -*-
import openpyxl
import csv
from smooth import Smoothener


class x2c():
    def _cell(self, xcell, smooth):
        value = smooth.celling(str(xcell.value))
        if xcell.data_type == xcell.TYPE_NUMERIC and xcell.number_format == '"0"0':  # Sonderfall #1 Telefonnummern
            value = "0" + value
        return value

    def convert_to_csv(self, load_file: str, save_file: str):
        print("Öffne XLSX: {}\nDies kann je nach Dateigröße ein paar Minuten dauern...".format(load_file))
        wb = openpyxl.load_workbook(load_file)
        sh = wb.active

        print("Schreibe CSV: {}".format(save_file))
        smooth = Smoothener()  # class die die Zellen bereinigt
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
                c.writerow(
                    [self._cell(cell, smooth) for cell in r])

        wb.close()

        if smooth.encoding_solved > 0:
            print(f"Es wurden {smooth.encoding_solved} Encodierungsfehler behoben. Bitte beachten! ")
        if smooth.encoding_issue > 0:
            print(f"Es wurden {smooth.encoding_issue} Encodierungsfehler nicht gelöst! Bitte beachten! ")

        print("Die Datei kann nun in Salesforce importiert werden:")
        print("z.B. DataImporter: Import-Typ: 'CSV', Zeichencode 'UTF-8' und Werte getrennt: 'Registerkarte'")
        print("Achtung! Das Nachbearbeiten mit Excel kann die Datei zerstören!")
        print("done")




