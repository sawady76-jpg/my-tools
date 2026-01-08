import pandas as pd
import glob
import os
import sys

# --- 設定：出力ファイル名 ---
OUTPUT_FILENAME = "発着信履歴_統合版.xlsx"

def main():
    print("--- 処理開始 ---")

    # 1. 実行場所を、このファイルがあるフォルダに強制変更（ここが重要！）
    try:
        current_dir = os.path.dirname(os.path.abspath(__file__))
        os.chdir(current_dir)
        print(f"フォルダ位置を固定しました: {current_dir}")
    except Exception:
        pass # インタラクティブモード等で失敗した場合は無視

    # 2. Excelファイルを探す
    # (自分自身＝出力ファイル は読み込まないように除外)
    all_files = sorted([f for f in glob.glob("*.xlsx") if OUTPUT_FILENAME not in f])

    if not all_files:
        print("【エラー】Excelファイル(.xlsx)が見つかりません。")
        print("このスクリプトと同じフォルダにデータを入れてください。")
        return

    print(f"対象ファイル数: {len(all_files)} 個")

    # 統合の設定
    # 縦に積み上げるシート
    merge_sheets = ["内線通話", "外線発信", "外線着信"]
    # 最新の1つだけにするシート
    single_sheets = ["着信件数", "電話端末"]

    # データ格納用
    merge_store = {k: [] for k in merge_sheets}
    single_store = {k: None for k in single_sheets}

    # 3. ファイル読み込み
    for file_path in all_files:
        print(f"・読み込み中: {file_path}")
        
        try:
            # 全シート一括読み込み
            xls_data = pd.read_excel(file_path, sheet_name=None)

            for sheet_name, df in xls_data.items():
                # 統合対象ならリストへ
                if sheet_name in merge_sheets:
                    merge_store[sheet_name].append(df)
                
                # 単一対象なら上書き（ファイル名順なので最後が最新）
                elif sheet_name in single_sheets:
                    single_store[sheet_name] = df

        except Exception as e:
            print(f"  [!] 読み込みエラー: {e}")

    # 4. 書き出し
    print("\n統合ファイルを作成しています...")

    try:
        with pd.ExcelWriter(OUTPUT_FILENAME, engine='openpyxl') as writer:
            
            # 統合シート
            for sheet_name, df_list in merge_store.items():
                if df_list:
                    combined_df = pd.concat(df_list, ignore_index=True)
                    combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"  OK: {sheet_name} ({len(combined_df)}行)")
            
            # 単一シート
            for sheet_name, df in single_store.items():
                if df is not None:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"  OK: {sheet_name}")

        print("\n" + "="*30)
        print(f" 完了！ 『{OUTPUT_FILENAME}』 が作成されました。")
        print("="*30)

    except PermissionError:
        print("\n【エラー】ファイルが開かれています。")
        print(f"『{OUTPUT_FILENAME}』を閉じてから再実行してください。")

if __name__ == "__main__":
    main()