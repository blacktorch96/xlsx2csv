# -*- encoding: utf-8 -*-
import tkinter as tk
from tkinter import filedialog
from xls2csv import x2c
import os

root = tk.Tk()
root.withdraw()

load_file = filedialog.askopenfilename(title="Excel-Datei zum konvertieren auswählen",
                                       filetypes=(("xlsx files", "*.xlsx"), ("all files", "*.*")))
if load_file == "":
    print("Keine Datei angegeben. Abbruch.")
    exit()

file_root, file_ext = os.path.splitext(load_file)
save_file = file_root + ".csv"

print(load_file)
print(save_file)

xlsx2csv = x2c()
xlsx2csv.convert_to_csv(load_file, save_file)
