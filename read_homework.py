import pandas as pd
import sys

file_path = r"D:\Estadistica Inferencial\Actividad de seleccion de muestra.xlsx"

try:
    with open("homework_summary.txt", "w", encoding="utf-8") as f:
        xl = pd.ExcelFile(file_path)
        f.write(f"Sheets in the file: {xl.sheet_names}\n")
        
        for sheet in xl.sheet_names:
            f.write(f"\n--- Sheet: {sheet} ---\n")
            df = xl.parse(sheet)
            
            df.dropna(how='all', axis=1, inplace=True)
            df.dropna(how='all', axis=0, inplace=True)
            
            pd.set_option('display.max_columns', None)
            pd.set_option('display.max_colwidth', None)
            pd.set_option('display.width', 1000)
            
            f.write(df.head(30).to_string() + "\n")
            
except Exception as e:
    with open("homework_summary.txt", "w", encoding="utf-8") as f:
        f.write(f"Error: {e}\n")
