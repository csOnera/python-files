import pathlib
import openpyxl
import os

folderi = pathlib.Path(r"C:\Users\onera\OneDrive - ONE ERA (HK) LIMITED\oneraShare\得物對賬")

folders = list(folderi.iterdir())

for folder in folders:
    if os.path.isdir(folder):
        exceli = pathlib.Path(folder)
        excelFiles = list(exceli.iterdir())
        for excel in excelFiles:
            toxlsx = excel.name + '.xlsx'
            os.rename(excel, os.path.join(folder, toxlsx))