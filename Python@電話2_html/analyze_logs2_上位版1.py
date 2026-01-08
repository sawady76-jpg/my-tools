import pandas as pd
import glob
import re
import os
import sys
import datetime
import warnings

# 警告を無視
warnings.simplefilter('ignore')

# ==========================================
# 1. 設定エリア
# ==========================================
BASE_ORDER = [
    '東京', '横浜', '埼玉', '滋賀', '大阪', 
    '千葉', '福岡', '岡山', '名古屋', '仙台', '流山'
]
EXCLUDE_KEYWORDS = ['不在', '未応答', '応答なし', '放棄', '留守電']
BIZ_START = datetime.time(8, 45, 0)
BIZ_END = datetime.time(17, 45, 0)
TRANS_FROM = {
    '東京': '流山', '横浜': '多摩', '埼玉': '千葉・北関東',
    '滋賀': '金沢', '福岡': '広島・熊本', '岡山': '大阪・静岡', '仙台': '札幌'
}
TRANS_TO = {'千葉': '埼玉', '大阪': '岡山'}

# ★追加集計用のグループ定義
SUMMARY_GROUPS = [
    {'name': '東京+流山', 'members': ['東京', '流山']},
    {'name': '埼玉+千葉+北関東', 'members': ['埼玉', '千葉', '北関東']},
    {'name': '横浜+多摩', 'members': ['横浜', '多摩']},
    {'name': '仙台+札幌', 'members': ['仙台', '札幌']},
    {'name': '滋賀+金沢', 'members': ['滋賀', '金沢']},
    {'name': '名古屋', 'members': ['名古屋']},
    {'name': '福岡+熊本+広島', 'members': ['福岡', '熊本', '広島']},
    {'name': '岡山+大阪+静岡', 'members': ['岡山', '大阪', '静岡']}
]

# ==========================================
# 2. 関数定義
# ==========================================

def extract_base_name(text):
    if not isinstance(text, str): return None, None
    match = re.search(r'【(.*?)】(.*)', text)
    if match: return match.group(1), match.group(2).strip()
    return None, text

def is_valid_answer(name):
    if not isinstance(name, str): return False
    for kw in EXCLUDE_KEYWORDS:
        if kw in name: return False
    return True

def my_round(x):
    try:
        if pd.isna(x): return ""
        val = float(x)
        return int(val * 10 + 0.5) / 10.0
    except: return x

def smart_read_excel(file, sheet_name):
    try:
        preview = pd.read_excel(file, sheet_name=sheet_name, header=None, nrows=20)
        target_row = 0
        for i, row in preview.iterrows():
            row_str = row.astype(str).values
            if any('着信者' in s for s in row_str) or any('時刻' in s for s in row_str):
                target_row = i
                break
        return pd.read_excel(file, sheet_name=sheet_name, header=target_row)
    except:
        return pd.DataFrame()

def find_and_load(target_keywords, exclude_keywords=[]):
    candidates = []
    for file in glob.glob("*.xlsx"):
        if "集計結果" in file: continue
        try:
            xls = pd.ExcelFile(file)
            for sheet in xls.sheet_names:
                if any(k in sheet for k in target_keywords):
                    if not any(ek in sheet for ek in exclude_keywords):
                        candidates.append({'file': file, 'sheet': sheet, 'type': 'xlsx'})
        except: pass

    for file in glob.glob("*.csv"):
        if "集計結果" in file: continue
        if any(k in file for k in target_keywords):
            if not any(ek in file for ek in exclude_keywords):
                candidates.append({'file': file, 'sheet': None, 'type': 'csv'})

    if not candidates:
        return pd.DataFrame()

    candidates.sort(key=lambda x: (0 if "統合版" in x['file'] else 1))
    
    target = candidates[0]
    print(f"  -> 読み込み: {os.path.basename(target['file'])} (Sheet: {target['sheet'] if target['sheet'] else 'CSV'})")
    
    if target['type'] == 'xlsx':
        return smart_read_excel(target['file'], target['sheet'])
    else:
        try: return pd.read_csv(target['file'], encoding='utf-8')
        except: return pd.read_csv(target['file'], encoding='cp932')

# ==========================================
# 3. メイン処理
# ==========================================
def main():
    try: os.chdir(os.path.dirname(os.path.abspath(__file__)))
    except: pass
    
    print(f"作業フォルダ: {os.getcwd()}")
    print("データを探しています...")

    df_int = find_and_load(['内線通話', '内線'], exclude_keywords=['発信'])
    df_ext = find_and_load(['外線着信', '外線'], exclude_keywords=['発信'])
    df_short = find_and_load(['時短勤務', '時短'])

    print(f"\nデータ件数: 内線={len(df_int)}件, 外線={len(df_ext)}件, 時短={len(df_short)}件")

    if df_int.empty and df_ext.empty:
        print("\n【エラー】データが見つかりません。")
        return

    data_year_month = None
    for df in [df_int, df_ext]:
        if df.empty: continue
        df.columns = [str(c).strip() for c in df.columns]
        if '着信者' in df.columns:
            df['target_base'], _ = zip(*df['着信者'].apply(extract_base_name))
        else: df['target_base'] = None
        if '最終着信者' in df.columns:
            df['final_base'], df['final_name'] = zip(*df['最終着信者'].apply(extract_base_name))
        else:
            df['final_base'] = None
            df['final_name'] = None
        if '時刻' in df.columns:
            df['dt'] = pd.to_datetime(df['時刻'].astype(str), errors='coerce')
            df['day'] = df['dt'].dt.day
            df['time'] = df['dt'].dt.time
            df['date'] = df['dt'].dt.date
            if data_year_month is None and not df['dt'].dropna().empty:
                first_date = df['dt'].dropna().iloc[0]
                data_year_month = (first_date.year, first_date.month)

    all_dates = set()
    if not df_int.empty: all_dates.update(df_int['date'].dropna().unique())
    if not df_ext.empty: all_dates.update(df_ext['date'].dropna().unique())
    total_operating_days = len(all_dates) if all_dates else 1
    
    # 3. 集計
    # --- Sheet 1 (着信件数) ---
    if not df_int.empty and 'final_name' in df_int.columns:
        df_int_valid = df_int[df_int['final_name'].apply(is_valid_answer)].copy()
    else: df_int_valid = pd.DataFrame()

    if not df_ext.empty and 'final_name' in df_ext.columns:
        df_ext_valid = df_ext[df_ext['final_name'].apply(is_valid_answer)].copy()
    else: df_ext_valid = pd.DataFrame()

    all_bases = set(BASE_ORDER)
    if not df_int.empty and 'target_base' in df_int.columns: all_bases |= set(df_int['target_base'].dropna())
    if not df_ext.empty and 'target_base' in df_ext.columns: all_bases |= set(df_ext['target_base'].dropna())
    sorted_bases = sorted(list(all_bases), key=lambda x: BASE_ORDER.index(x) if x in BASE_ORDER else 999)

    s1 = pd.DataFrame(index=sorted_bases)
    s1.index.name = '拠点'
    
    def get_counts(df, col):
        if df.empty or col not in df.columns: return {}
        return df[col].value_counts()

    # Sheet1 本表
    s1_data = {
        ('内線', '入電'): s1.index.map(get_counts(df_int, 'target_base')).fillna(0).astype(int),
        ('内線', '着電'): s1.index.map(get_counts(df_int_valid, 'final_base')).fillna(0).astype(int),
        ('外線', '入電'): s1.index.map(get_counts(df_ext, 'target_base')).fillna(0).astype(int),
        ('外線', '着電'): s1.index.map(get_counts(df_ext_valid, 'final_base')).fillna(0).astype(int),
        ('他拠点へ転送', ''): s1.index.map(TRANS_TO).fillna(''),
        ('他拠点から転送', ''): s1.index.map(TRANS_FROM).fillna('')
    }
    s1 = pd.DataFrame(s1_data, index=s1.index)

    # ---------------------------------------------------------
    # ★J12集計エリア: グループ別合算集計（外線のみ対象）
    # ---------------------------------------------------------
    summary_rows = [['追加集計', '電話が入った数', 'とった数', '％']]
    
    for grp in SUMMARY_GROUPS:
        # グループに属する拠点の数値を合算
        # 存在しない拠点は無視する
        members = [m for m in grp['members'] if m in s1.index]
        
        if members:
            # 外線の入電・着電を合計
            sum_in = s1.loc[members, ('外線', '入電')].sum()
            sum_ans = s1.loc[members, ('外線', '着電')].sum()
            rate = f"{sum_ans / sum_in:.1%}" if sum_in > 0 else "0.0%"
            
            summary_rows.append([grp['name'], sum_in, sum_ans, rate])
        else:
            # メンバーがデータに存在しない場合
            summary_rows.append([grp['name'], 0, 0, "0.0%"])

    s1_summary = pd.DataFrame(summary_rows)
    # ---------------------------------------------------------

    # --- Sheet 2 (従業員別) ---
    cols = ['final_name', 'final_base', '通話時間', 'day', 'date']
    s2 = pd.DataFrame()

    if (not df_int_valid.empty) or (not df_ext_valid.empty):
        if not df_int_valid.empty and set(cols).issubset(df_int_valid.columns):
            grp_int = df_int_valid.groupby(['final_name', 'final_base'])
            s2_int = grp_int.agg(内線件数=('通話時間', 'size'), 内線合計=('通話時間', 'sum'))
        else: s2_int = pd.DataFrame(columns=['内線件数', '内線合計'])

        if not df_ext_valid.empty and set(cols).issubset(df_ext_valid.columns):
            grp_ext = df_ext_valid.groupby(['final_name', 'final_base'])
            s2_ext = grp_ext.agg(外線件数=('通話時間', 'size'), 外線合計=('通話時間', 'sum'))
        else: s2_ext = pd.DataFrame(columns=['外線件数', '外線合計'])

        s2 = s2_int.join(s2_ext, how='outer').fillna(0)
        s2['受発注'] = ""
        s2['内線'] = s2['内線件数']
        s2['通話時間／秒'] = (s2['内線合計'] / s2['内線件数'].replace(0, 1)).apply(my_round)
        s2['外線'] = s2['外線件数']
        s2['通話時間／秒.1'] = (s2['外線合計'] / s2['外線件数'].replace(0, 1)).apply(my_round)
        
        df_all_valid = pd.concat([df_int_valid[cols] if not df_int_valid.empty else pd.DataFrame(), df_ext_valid[cols] if not df_ext_valid.empty else pd.DataFrame()])
        
        if not df_all_valid.empty:
            if data_year_month:
                y, m = data_year_month
                daily = df_all_valid.pivot_table(index=['final_name', 'final_base'], columns='day', aggfunc='size', fill_value=0)
                date_cols = {}
                for d in range(1, 32):
                    try:
                        date_str = datetime.date(y, m, d).strftime('%Y-%m-%d')
                        if d not in daily.columns: daily[d] = 0
                        date_cols[d] = date_str
                    except ValueError: pass
                daily = daily.rename(columns=date_cols)
                daily = daily[sorted(list(date_cols.values()))]
                daily['稼働日'] = (daily > 0).sum(axis=1)
                s2 = s2.join(daily).reset_index()
                s2['内外線計'] = s2['内線'] + s2['外線']
                s2['1日平均'] = s2.apply(lambda r: my_round(r['内外線計'] / r['稼働日']) if r['稼働日'] > 0 else 0, axis=1)
                s2 = s2.rename(columns={'final_name': '名前', 'final_base': '拠点'})
                s2['base_rank'] = s2['拠点'].apply(lambda x: BASE_ORDER.index(x) if x in BASE_ORDER else 999)
                s2 = s2.sort_values(['base_rank', '名前']).drop(columns='base_rank')
                final_cols = ['名前', '拠点', '受発注', '内線', '通話時間／秒', '外線', '通話時間／秒.1'] + sorted(list(date_cols.values())) + ['稼働日', '内外線計', '1日平均']
                s2 = s2[[c for c in final_cols if c in s2.columns]]

    # --- Sheet 3 (拠点別) ---
    if ('外線', '着電') in s1.columns:
        s3 = s1[[('外線', '着電')]].reset_index().copy()
        s3.columns = ['拠点名', '外線(実績)']
    else: s3 = pd.DataFrame(columns=['拠点名', '外線(実績)'])

    col_name_total = f"{data_year_month[0]}年{data_year_month[1]}月から外線のみ" if data_year_month else "期間計"
    s3 = s3.rename(columns={'外線(実績)': col_name_total})
    s3['外線のみ'] = s3[col_name_total].apply(lambda x: my_round(x / total_operating_days) if total_operating_days > 0 else 0)

    if not s2.empty:
        personnel = s2.groupby('拠点')['名前'].nunique()
        s3['人員'] = s3['拠点名'].map(personnel).fillna(0).astype(int)
    else: s3['人員'] = 0

    s3['1人当たり／月'] = s3.apply(lambda r: my_round(r[col_name_total] / r['人員']) if r['人員'] > 0 else 0.0, axis=1)
    total_s3 = s3[col_name_total].sum()
    s3['全体からの比率'] = s3[col_name_total].apply(lambda x: f"{my_round(x / total_s3 * 100)}%" if total_s3 > 0 else "0.0%")
    s3 = s3[['拠点名', col_name_total, '外線のみ', '人員', '1人当たり／月', '全体からの比率']]

    # --- Sheet 4 (時短勤務) ---
    if not df_short.empty:
        for col in df_short.select_dtypes(include='number').columns:
            df_short[col] = df_short[col].apply(my_round)
    
    # --- Sheet 5 (営業時間内集計) ---
    s5 = pd.DataFrame(index=sorted_bases)
    s5.index.name = '拠点名'
    if not df_ext_valid.empty and 'time' in df_ext_valid.columns:
        df_ext_biz = df_ext_valid[df_ext_valid['time'].apply(lambda t: BIZ_START <= t <= BIZ_END)]
        s5['営業時間内_外線のみ'] = s5.index.map(df_ext_biz['final_base'].value_counts()).fillna(0).astype(int)
    else: s5['営業時間内_外線のみ'] = 0

    s5['人員'] = s5.index.map(s3.set_index('拠点名')['人員']).fillna(0).astype(int)
    s5['1人当たり／月'] = s5.apply(lambda r: my_round(r['営業時間内_外線のみ'] / r['人員']) if r['人員'] > 0 else 0.0, axis=1)
    total_s5 = s5['営業時間内_外線のみ'].sum()
    s5['全体からの比率'] = s5['営業時間内_外線のみ'].apply(lambda x: f"{my_round(x / total_s5 * 100)}%" if total_s5 > 0 else "0.0%")
    s5 = s5.reset_index()

    # 4. 出力
    print("\n集計中...")
    output_filename = f"集計結果_{datetime.date.today()}.xlsx"
    
    try:
        with pd.ExcelWriter(output_filename) as writer:
            # Sheet1
            s1.reset_index().to_excel(writer, sheet_name='1.着信件数', startrow=0, startcol=0)
            # ★J12 (index=11, col=9) に追加集計を出力
            s1_summary.to_excel(writer, sheet_name='1.着信件数', index=False, header=False, startrow=11, startcol=9)

            s2.to_excel(writer, sheet_name='2.従業員別', index=False)
            s3.to_excel(writer, sheet_name='3.関数_拠点別', index=False)
            if not df_short.empty: df_short.to_excel(writer, sheet_name='4.時短勤務', index=False)
            else: pd.DataFrame({'info': ['データなし']}).to_excel(writer, sheet_name='4.時短勤務', index=False)
            s5.to_excel(writer, sheet_name='5.営業時間内集計', index=False)

        print(f"\n★★ 完了しました！ ★★")
        print(f"作成されたファイル: {output_filename}")
        print("Sheet1のJ12に「グループ合算集計」を追加しました。")
    except Exception as e:
        print(f"\n【エラー】Excelファイルの保存に失敗しました: {e}")

    input("\nEnterキーを押して終了してください...")

if __name__ == "__main__":
    main()