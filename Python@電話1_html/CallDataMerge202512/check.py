import sys
import os
import glob

print("--- 診断開始 ---")

# 1. 実行場所の確認
current_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(current_dir)
print(f"現在のフォルダ: {current_dir}")

# 2. ライブラリの確認
print("ライブラリ確認中...", end="")
try:
    import pandas as pd
    import openpyxl
    print("OK (pandas, openpyxl はインストールされています)")
except ImportError as e:
    print("\n【エラー！】ライブラリが見つかりません。")
    print(f"詳細: {e}")
    print("対策: 黒い画面(コマンドプロンプト)で以下を実行してください:")
    print("pip install pandas openpyxl")
    input("Enterキーを押して終了...")
    sys.exit()

# 3. ファイルの確認
xlsx_files = glob.glob("*.xlsx")
print(f"\nExcelファイル検出数: {len(xlsx_files)} 個")

if len(xlsx_files) == 0:
    print("【エラー！】同じフォルダに .xlsx ファイルがありません。")
    print("・拡張子が .csv や .xls になっていませんか？")
    print("・このプログラムをExcelファイルと同じフォルダに入れましたか？")
else:
    print("見つかったファイル(最初の3つ):", xlsx_files[:3])
    
    # 4. 中身（シート名）の確認
    target_file = xlsx_files[0]
    print(f"\n『{target_file}』の中身をチェックします...")
    try:
        xls = pd.read_excel(target_file, sheet_name=None)
        print("含まれているシート名一覧:")
        for sheet in xls.keys():
            print(f" - {sheet}")
            
        print("\n--- 診断完了 ---")
        print("このシート名が、プログラムで指定している名前（内線通話、外線着信など）と")
        print("完全に一致しているか確認してください（スペースの有無など）。")
        
    except Exception as e:
        print(f"\n【エラー！】ファイルを読み込めませんでした。")
        print(f"詳細: {e}")

input("\nEnterキーを押して終了してください...")