# -*- coding: utf-8 -*-
import pandas as pd

xl = pd.ExcelFile('発着信履歴_統合版.xlsx')

with open('encoding_check_result.txt', 'w', encoding='utf-8') as f:
    f.write("=== シート名 ===\n")
    for i, name in enumerate(xl.sheet_names):
        f.write(f"  [{i}] {repr(name)}\n")
    f.write("\n")
    
    for sheet in xl.sheet_names:
        f.write(f"=== シート: {sheet} の列名 ===\n")
        df = pd.read_excel(xl, sheet_name=sheet, nrows=2)
        for col in df.columns:
            f.write(f"  {repr(col)}\n")
        f.write("\n")
        f.write("--- 先頭2行のデータ ---\n")
        f.write(df.to_string())
        f.write("\n\n")

print("Done! Check encoding_check_result.txt")
