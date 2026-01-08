# -*- coding: utf-8 -*-
import pandas as pd

xl = pd.ExcelFile('発着信履歴_統合版.xlsx')

with open('data_info.txt', 'w', encoding='utf-8') as f:
    f.write(f'シート名: {xl.sheet_names}\n\n')
    
    for sheet in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sheet)
        f.write(f'=== {sheet} ===\n')
        f.write(f'列名: {df.columns.tolist()}\n')
        f.write(f'行数: {len(df)}\n')
        f.write('最初の3行:\n')
        f.write(df.head(3).to_string())
        f.write('\n\n')

print('data_info.txt に出力しました')
