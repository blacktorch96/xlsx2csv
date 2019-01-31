# -*- encoding: utf-8 -*-


class Smoothener():
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

    def celling(self, value) -> str:
        # Ersetzt Zeilenumbrüche und Tabluatoren durch Leerzeichen
        value = value.replace('\n', ' ').replace('\t', ' ').replace('\r', ' ').replace("None", '')

        # Sind Sonderzeichen enthalten, dann die Zeichenliste durchgehen...
        if chr(0xC3) in value or chr(0xC2) in value or chr(0xC5) in value or chr(0xE2) in value or chr(
                0xCB) in value or chr(0xC6) in value:
            for k, v in self.enc.items():
                if k in value:
                    value = value.replace(k, v)
                    self.encoding_solved += 1
            if 'Ã' in value:
                self.encoding_issue += 1
                # todo: sind noch weitere Endcodierungsfehler vorhanden?
                print("weiterer Encodierungsfehler!")
        return value
