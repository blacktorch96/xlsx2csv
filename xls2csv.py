# -*- encoding: utf-8 -*-
import openpyxl
import csv


class x2c():

    def __init__(self):
        self.enc = {'â‚¬': '€', 'â€š': '‚', 'â€ž': '„', 'â€¦': '…', 'â€¡': '‡', 'â€°': '‰', 'â€¹': '‹', 'â€˜': '‘',
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
        self.encoding_issue = 0
        self.encoding_solved = 0

    def _celling(self, xcell) -> str:
        value = str(xcell.value).replace('\n', ' ').replace('\t', ' ').replace('\r', ' ').replace("None", '')
        if chr(0xC3) in value or chr(0xC2) in value or chr(0xC5) in value or chr(0xE2) in value or chr(0xCB) in value or chr(0xC6) in value:

            for k, v in self.enc.items():
                if k in value:
                    value = value.replace(k, v)
                    self.encoding_solved += 1
            if 'Ã' in value:
                self.encoding_issue += 1
                # todo: sind noch weitere Endcodierungsfehler vorhanden?
                print("weiterer Encodierungsfehler!")
        if xcell.data_type == xcell.TYPE_NUMERIC and xcell.number_format == '"0"0':  # Sonderfall #1 Telefonnummern
            value = "0" + value
        return value

    def convert_to_csv(self, load_file: str, save_file: str):
        print("Öffne XLSX: {}\nDies kann je nach Dateigröße ein paar Minuten dauern...".format(load_file))
        wb = openpyxl.load_workbook(load_file)
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
                c.writerow(
                    [self._celling(cell) for cell in r])

        wb.close()

        if self.encoding_solved > 0:
            print(f"Es wurden {self.encoding_solved} Encodierungsfehler behoben. Bitte beachten! ")
        if self.encoding_issue > 0:
            print(f"Es wurden {self.encoding_issue} Encodierungsfehler nicht gelöst! Bitte beachten! ")

        print("Die Datei kann nun in Salesforce importiert werden:")
        print("z.B. DataImporter: Import-Typ: 'CSV', Zeichencode 'UTF-8' und Werte getrennt: 'Registerkarte'")
        print("Achtung! Das Nachbearbeiten mit Excel kann die Datei zerstören!")
        print("done")




