import pandas as pd
import glob
import os
import sys

# --- 設定：出力ファイル名 ---
OUTPUT_FILENAME = "発着信履歴_統合版.xlsx"

# --- 設定：見本ファイルから取得した「正しいヘッダー」 ---
# 内線通話の元データはヘッダーが崩れているため、このリストで強制的に上書きします
INTERNAL_HEADERS = [
    '時刻', 
    '発信番号', 
    '発信者',          # 元データでは「不在」となっていた列
    '最終着信者名',    # 元データでは「着信番号」となっていた列
    '着信者',          # 元データでは「不在.1」となっていた列
    '最終着信番号', 
    '最終着信者',      # 元データでは「不在.2」となっていた列
    '通話時間（応答までの時間を含む）', 
    '通話時間', 
    'メモ', 
    'リクエストID'
]

def main():
    print("--- 処理開始 ---")

    # 1. 実行場所の固定
    try:
        os.chdir(os.path.dirname(os.path.abspath(__file__)))
    except:
        pass

    # 2. Excelファイルを探す (出力ファイルと見本ファイルは除外)
    all_files = sorted([f for f in glob.glob("*.xlsx") 
                        if OUTPUT_FILENAME not in f and "見本" not in f])

    if not all_files:
        print("【エラー】Excelファイル(.xlsx)が見つかりません。")
        input("Enterキーを押して終了してください...")
        return

    print(f"対象ファイル数: {len(all_files)} 個")

    # データ格納用リスト
    store = {
        "内線通話": [],
        "外線発信": [],
        "外線着信": []
    }

    # 3. 各ファイルを読み込み
    for file_path in all_files:
        print(f"・読み込み中: {file_path}")
        
        try:
            # 全シート読み込み
            xls_data = pd.read_excel(file_path, sheet_name=None)

            # --- シートごとの処理 ---
            
            # (1) 内線通話：ヘッダーを修正して取り込む
            if "内線通話" in xls_data:
                df = xls_data["内線通話"]
                # 列数が合っているか確認
                if len(df.columns) == len(INTERNAL_HEADERS):
                    # ヘッダーを強制上書き
                    df.columns = INTERNAL_HEADERS
                    store["内線通話"].append(df)
                else:
                    print(f"  [!] 注意: {file_path} の「内線通話」は列数が違うためスキップしました")

            # (2) 外線発信：そのまま取り込む
            if "外線発信" in xls_data:
                store["外線発信"].append(xls_data["外線発信"])

            # (3) 外線着信：そのまま取り込む
            if "外線着信" in xls_data:
                store["外線着信"].append(xls_data["外線着信"])

        except Exception as e:
            print(f"  [!] エラー: {e}")

    # 4. 統合して保存
    print("\n統合ファイルを作成しています...")

    try:
        with pd.ExcelWriter(OUTPUT_FILENAME, engine='openpyxl') as writer:
            for sheet_name, df_list in store.items():
                if df_list:
                    # 縦に結合
                    combined_df = pd.concat(df_list, ignore_index=True)
                    
                    # 見本通り、時刻順に並べ替えたい場合は以下を有効化してください
                    # if '時刻' in combined_df.columns:
                    #     combined_df = combined_df.sort_values('時刻')

                    combined_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    print(f"  OK: {sheet_name} ({len(combined_df)}行)")
                else:
                    print(f"  SKIP: {sheet_name} (データなし)")

        print("\n" + "="*30)
        print(f" 完了！ 『{OUTPUT_FILENAME}』 が作成されました。")
        print("="*30)

    except PermissionError:
        print("\n【エラー】ファイルが開かれています。閉じてから再実行してください。")

    input("Enterキーを押して終了してください...")

if __name__ == "__main__":
    main()