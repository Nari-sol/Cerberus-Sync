import streamlit as st
import pandas as pd
import io
import re
import plotly.graph_objects as go
import operator

# 注意: Excelの読み込みには 'openpyxl' ライブラリが必要です。
# インストールされていない場合は: pip install openpyxl を実行してください。

# ページ設定
st.set_page_config(
    page_title="Cerberus Sync",
    page_icon="🐕",
    layout="wide",
    initial_sidebar_state="expanded"
)

# カスタムCSS
st.markdown("""
    <style>
    .main { background-color: #0e1117; }
    .stTitle { color: #00f2fe; font-family: 'Inter', sans-serif; font-weight: 800; text-shadow: 0px 0px 10px rgba(0, 242, 254, 0.5); }
    .stSubheader { color: #4facfe; font-family: 'Inter', sans-serif; }
    </style>
    """, unsafe_allow_html=True)

# タイトル表示
st.title("Cerberus Sync")
st.subheader("マルチモール統合分析（Yahoo! & 楽天）")

# 対象月を抽出するヘルパー関数
def extract_month_from_str(s):
    numbers = re.findall(r'\d+', s)
    for num in numbers:
        val = int(num)
        if 1 <= val <= 12:
            return f"{val}月"
    return "不明"

# 店舗名の判定ロジック
def get_store_name(filename):
    if "楽天" in filename:
        return "楽天"
    elif "YS1" in filename: return "Yahoo1号店"
    elif "YS2" in filename: return "Yahoo2号店"
    elif "YS3" in filename: return "Yahoo3号店"
    elif "YS4" in filename: return "Yahoo4号店"
    elif "YS5" in filename: return "Yahoo5号店"
    elif "Amazon1" in filename or "UBUNBASE_1" in filename or "UBUNBASE1" in filename: return "Amazon1号店"
    elif "Amazon2" in filename or "UBUNBASE_2" in filename or "UBUNBASE2" in filename: return "Amazon2号店"
    else: return f"不明_{filename.replace('.xlsx', '')}"

# データ読み込み・クレンジング・結合処理のキャッシュ化
@st.cache_data
def load_and_process_data(uploaded_files_data, ad_uploaded_file_data, yahoo_ad_uploaded_files_data, yahoo_item_uploaded_files_data):
    warnings_list = []
    errors_list = []
    all_data_list = []
    yahoo_item_dfs = []
    amazon_budget_loaded_months = set()
    amazon_extracted_dfs = []

    # 楽天RPP広告データの読み込み・クレンジング
    ad_df = None
    if ad_uploaded_file_data:
        ad_name, ad_content = ad_uploaded_file_data
        try:
            # 拡張子に応じて読み込み処理を分岐し、全シートを結合
            if ad_name.lower().endswith('.xlsx'):
                sheets_dict = pd.read_excel(io.BytesIO(ad_content), sheet_name=None)
                sheet_names = list(sheets_dict.keys())
                
                # 割引データシートの分離（最初のシート）
                discount_sheet = sheet_names[0]
                discount_df = sheets_dict[discount_sheet]
                
                # 月と割引金額のマッピングを作成
                discount_map = {}
                if len(discount_df.columns) >= 2:
                    for _, row in discount_df.iterrows():
                        m_val = str(row.iloc[0]).strip()
                        m_key = extract_month_from_str(m_val)
                        try:
                            amt_str = str(row.iloc[1]).replace(',', '').replace('円', '').replace('¥', '').strip()
                            amt = float(amt_str) if amt_str else 0.0
                        except Exception:
                            amt = 0.0
                        
                        discount_map[m_val] = amt
                        if m_key != "不明":
                            discount_map[m_key] = amt
                
                # 2番目以降のシート（各月の広告データ）をループ処理
                cleaned_sheets = []
                for sname in sheet_names[1:]:
                    df = sheets_dict[sname].copy()
                    if df.empty:
                        continue
                    
                    df.columns = df.columns.astype(str).str.replace(r'[\s　]+', '', regex=True)
                    
                    # 割引金額の取得
                    m_key = extract_month_from_str(sname)
                    discount_amt = discount_map.get(sname, 0.0)
                    if discount_amt == 0.0 and m_key != "不明":
                        discount_amt = discount_map.get(m_key, 0.0)
                    
                    # 日付列と広告費列の特定
                    date_col = None
                    for col in ["日付", "日", "年月日"]:
                        if col in df.columns:
                            date_col = col
                            break
                    
                    cost_col = None
                    for col in ["実績額(合計)", "割引後実績額", "実績額", "広告費", "利用金額", "利用額"]:
                        if col in df.columns:
                            cost_col = col
                            break
                    
                    # 割引額の反映
                    if date_col:
                        if pd.api.types.is_numeric_dtype(df[date_col]):
                            df[date_col] = pd.to_datetime(df[date_col], unit='D', origin='1899-12-30')
                        else:
                            df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
                        df[date_col] = df[date_col].dt.normalize()
                        
                        # ユニーク日付数のカウント
                        unique_dates = df[date_col].dropna().unique()
                        num_days = len(unique_dates)
                        
                        if num_days > 0 and discount_amt > 0.0:
                            daily_discount = discount_amt / num_days
                            
                            if cost_col:
                                df[cost_col] = pd.to_numeric(
                                    df[cost_col].astype(str).str.replace(',', '', regex=False).str.replace('円', '', regex=False).str.replace('¥', '', regex=False).str.strip(),
                                    errors='coerce'
                                ).fillna(0.0)
                                df[cost_col] = df[cost_col].apply(lambda val: max(val - daily_discount, 0.0))
                    
                    cleaned_sheets.append(df)
                
                if cleaned_sheets:
                    ads_df = pd.concat(cleaned_sheets, ignore_index=True)
                else:
                    ads_df = pd.DataFrame()
            else:
                try:
                    ads_df = pd.read_csv(io.BytesIO(ad_content), encoding='cp932')
                except Exception:
                    ads_df = pd.read_csv(io.BytesIO(ad_content), encoding='utf-8')
                ads_df.columns = ads_df.columns.astype(str).str.replace(r'[\s　]+', '', regex=True)
            
            cost_col = None
            for col in ["実績額(合計)", "割引後実績額", "実績額", "広告費", "利用金額", "利用額"]:
                if col in ads_df.columns:
                    cost_col = col
                    break
            
            sales_col = None
            for col in ["売上金額(合計720時間)", "売上金額(合計)", "売上金額(合計12時間)", "売上金額", "売上額", "広告経由売上", "広告売上"]:
                if col in ads_df.columns:
                    sales_col = col
                    break
            
            click_col = None
            for col in ["クリック数(合計)", "クリック数", "クリック"]:
                if col in ads_df.columns:
                    click_col = col
                    break
            
            orders_col = None
            for col in ["売上件数(合計720時間)", "売上件数(合計)", "売上件数", "注文数", "件数"]:
                if col in ads_df.columns:
                    orders_col = col
                    break
            
            date_col = None
            for col in ["日付", "日", "年月日"]:
                if col in ads_df.columns:
                    date_col = col
                    break
            
            # 日付列の変換処理（Excelシリアル値対応）
            if date_col:
                if pd.api.types.is_numeric_dtype(ads_df[date_col]):
                    ads_df[date_col] = pd.to_datetime(ads_df[date_col], unit='D', origin='1899-12-30')
                else:
                    ads_df[date_col] = pd.to_datetime(ads_df[date_col], errors='coerce')
                ads_df[date_col] = ads_df[date_col].dt.normalize()
            
            if date_col and cost_col and sales_col:
                ad_temp = pd.DataFrame()
                ad_temp["日付"] = ads_df[date_col]
                
                ad_temp["広告費"] = pd.to_numeric(
                    ads_df[cost_col].astype(str).str.replace(',', '', regex=False).str.replace('円', '', regex=False).str.replace('¥', '', regex=False).str.strip(), 
                    errors='coerce'
                ).fillna(0)
                
                ad_temp["広告売上"] = pd.to_numeric(
                    ads_df[sales_col].astype(str).str.replace(',', '', regex=False).str.replace('円', '', regex=False).str.replace('¥', '', regex=False).str.strip(), 
                    errors='coerce'
                ).fillna(0)
                
                if click_col:
                    ad_temp["クリック数"] = pd.to_numeric(
                        ads_df[click_col].astype(str).str.replace(',', '', regex=False), 
                        errors='coerce'
                    ).fillna(0)
                else:
                    ad_temp["クリック数"] = 0.0
                    
                if orders_col:
                    ad_temp["注文数"] = pd.to_numeric(
                        ads_df[orders_col].astype(str).str.replace(',', '', regex=False), 
                        errors='coerce'
                    ).fillna(0)
                else:
                    ad_temp["注文数"] = 0.0
                
                # 同一日の複数行データをマージ（日単位で集計）
                ad_df = ad_temp.groupby("日付")[["広告費", "広告売上", "クリック数", "注文数"]].sum().reset_index()
            else:
                missing_cols = []
                if not date_col: missing_cols.append("日付")
                if not cost_col: missing_cols.append("実績額(合計)などの広告費")
                if not sales_col: missing_cols.append("売上金額(合計720時間)などの広告売上")
                
                warnings_list.append(
                    f"⚠️ 楽天RPP広告レポートの必要な列が見つかりませんでした。\n"
                    f"不足していると判定された列: {', '.join(missing_cols)}\n\n"
                    f"実際に入力された列名リスト: {ads_df.columns.tolist()}"
                )
        except Exception as e:
            errors_list.append(
                f"❌ 広告レポートの読み込みまたは計算中にエラーが発生しました: {e}\n\n"
                f"読み込んだデータの列名リスト: {ads_df.columns.tolist() if 'ads_df' in locals() else 'ファイル読み込み失敗'}"
            )
            
    # Yahoo!アイテムマッチ広告データの読み込み・クレンジング（複数ファイル対応・店舗自動仕分け）
    yahoo_ad_df = None
    if yahoo_ad_uploaded_files_data:
        y_dfs = []
        for y_name, y_content in yahoo_ad_uploaded_files_data:
            try:
                store_label = "Yahoo!"
                if "YS1" in y_name:
                    store_label = "Yahoo1号店"
                elif "YS2" in y_name:
                    store_label = "Yahoo2号店"
                
                if y_name.lower().endswith('.xlsx'):
                    y_sheets_dict = pd.read_excel(io.BytesIO(y_content), sheet_name=None)
                    y_cleaned_sheets = []
                    for sname, df in y_sheets_dict.items():
                        df.columns = df.columns.astype(str).str.replace(r'[\s　]+', '', regex=True)
                        y_cleaned_sheets.append(df)
                    yahoo_ads_df = pd.concat(y_cleaned_sheets, ignore_index=True)
                else:
                    try:
                        yahoo_ads_df = pd.read_csv(io.BytesIO(y_content), encoding='cp932')
                    except Exception:
                        yahoo_ads_df = pd.read_csv(io.BytesIO(y_content), encoding='utf-8')
                    yahoo_ads_df.columns = yahoo_ads_df.columns.astype(str).str.replace(r'[\s　]+', '', regex=True)
                
                y_cost_col = None
                for col in ["利用金額(請求額)", "利用金額", "利用額", "実績額", "広告費"]:
                    if col in yahoo_ads_df.columns:
                        y_cost_col = col
                        break
                
                y_sales_col = None
                for col in ["売上金額", "売上額", "売上", "広告売上", "広告経由売上"]:
                    if col in yahoo_ads_df.columns:
                        y_sales_col = col
                        break
                
                y_click_col = None
                for col in ["クリック数", "クリック"]:
                    if col in yahoo_ads_df.columns:
                        y_click_col = col
                        break
                
                y_orders_col = None
                for col in ["注文数", "注件数", "件数"]:
                    if col in yahoo_ads_df.columns:
                        y_orders_col = col
                        break
                
                y_date_col = None
                for col in ["日付", "日", "年月日"]:
                    if col in yahoo_ads_df.columns:
                        y_date_col = col
                        break
                
                # 日付列の変換処理（Excelシリアル値対応）
                if y_date_col:
                    if pd.api.types.is_numeric_dtype(yahoo_ads_df[y_date_col]):
                        yahoo_ads_df[y_date_col] = pd.to_datetime(yahoo_ads_df[y_date_col], unit='D', origin='1899-12-30')
                    else:
                        yahoo_ads_df[y_date_col] = pd.to_datetime(yahoo_ads_df[y_date_col], errors='coerce')
                    yahoo_ads_df[y_date_col] = yahoo_ads_df[y_date_col].dt.normalize()
                
                if y_date_col and y_cost_col and y_sales_col:
                    y_ad_temp = pd.DataFrame()
                    y_ad_temp["日付"] = yahoo_ads_df[y_date_col]
                    
                    y_ad_temp["広告費"] = pd.to_numeric(
                        yahoo_ads_df[y_cost_col].astype(str).str.replace(',', '', regex=False).str.replace('円', '', regex=False).str.replace('¥', '', regex=False).str.strip(), 
                        errors='coerce'
                    ).fillna(0)
                    
                    y_ad_temp["広告売上"] = pd.to_numeric(
                        yahoo_ads_df[y_sales_col].astype(str).str.replace(',', '', regex=False).str.replace('円', '', regex=False).str.replace('¥', '', regex=False).str.strip(), 
                        errors='coerce'
                    ).fillna(0)
                    
                    if y_click_col:
                        y_ad_temp["クリック数"] = pd.to_numeric(
                            yahoo_ads_df[y_click_col].astype(str).str.replace(',', '', regex=False),
                            errors='coerce'
                        ).fillna(0)
                    else:
                        y_ad_temp["クリック数"] = 0.0
                        
                    if y_orders_col:
                        y_ad_temp["注文数"] = pd.to_numeric(
                            yahoo_ads_df[y_orders_col].astype(str).str.replace(',', '', regex=False),
                            errors='coerce'
                        ).fillna(0)
                    else:
                        y_ad_temp["注文数"] = 0.0
                    
                    y_ad_temp["店舗"] = store_label
                    y_dfs.append(y_ad_temp)
                else:
                    y_missing_cols = []
                    if not y_date_col: y_missing_cols.append("日付")
                    if not y_cost_col: y_missing_cols.append("利用金額(請求額)などの広告費")
                    if not y_sales_col: y_missing_cols.append("売上金額などの広告売上")
                    
                    warnings_list.append(
                        f"⚠️ Yahoo!アイテムマッチ広告レポート（{y_name}）の必要な列が見つかりませんでした。\n"
                        f"不足していると判定された列: {', '.join(y_missing_cols)}"
                    )
            except Exception as e:
                errors_list.append(f"❌ Yahoo!広告レポート（{y_name}）の読み込みエラー: {e}")
        
        if y_dfs:
            yahoo_ads_df_all = pd.concat(y_dfs, ignore_index=True)
            yahoo_ad_df = yahoo_ads_df_all.groupby(["日付", "店舗"])[["広告費", "広告売上", "クリック数", "注文数"]].sum().reset_index()
            
    # 商品別データの読み込み・結合
    if yahoo_item_uploaded_files_data:
        for y_item_name, y_item_content in yahoo_item_uploaded_files_data:
            try:
                df_item = None
                if "楽天" in y_item_name:
                    if y_item_name.lower().endswith('.xlsx'):
                        sheets_dict = pd.read_excel(io.BytesIO(y_item_content), sheet_name=None)
                        r_dfs = []
                        for sheet_name, df_sheet in sheets_dict.items():
                            df_sheet.columns = [str(c).strip() for c in df_sheet.columns]
                            df_sheet["対象月"] = extract_month_from_str(sheet_name)
                            r_dfs.append(df_sheet)
                        df_item = pd.concat(r_dfs, ignore_index=True)
                    else:
                        try:
                            df_item = pd.read_csv(io.BytesIO(y_item_content), encoding='cp932')
                        except Exception:
                            df_item = pd.read_csv(io.BytesIO(y_item_content), encoding='utf-8')
                        df_item.columns = [str(c).strip() for c in df_item.columns]
                        df_item["対象月"] = extract_month_from_str(y_item_name)
                    
                    if "商品管理番号" in df_item.columns and "商品番号" in df_item.columns:
                        df_item = df_item.rename(columns={"商品番号": "元_商品番号"}, errors='ignore')
                    
                    df_item = df_item.rename(columns={
                        "商品管理番号": "商品コード",
                        "商品番号": "商品コード",
                        "売上": "売上合計値（税込）",
                        "売上個数": "注文点数"
                    }, errors='ignore')
                    df_item = df_item.loc[:, ~df_item.columns.duplicated()]
                    if "商品コード" in df_item.columns:
                        df_item["商品コード"] = df_item["商品コード"].fillna("不明").astype(str)
                    
                    df_item["店舗名"] = "楽天"
                    yahoo_item_dfs.append(df_item)
                elif "Amazon1" in y_item_name or "Amazon2" in y_item_name:
                    item_store_name = "Amazon1号店" if "Amazon1" in y_item_name else "Amazon2号店"
                    if y_item_name.lower().endswith('.xlsx'):
                        sheets_dict = pd.read_excel(io.BytesIO(y_item_content), sheet_name=None)
                        a_dfs = []
                        for sheet_name, df_sheet in sheets_dict.items():
                            df_sheet.columns = [str(c).strip() for c in df_sheet.columns]
                            m_val = extract_month_from_str(sheet_name)
                            if m_val == "不明":
                                m_val = extract_month_from_str(y_item_name)
                            df_sheet["対象月"] = m_val
                            a_dfs.append(df_sheet)
                        df_item = pd.concat(a_dfs, ignore_index=True)
                    else:
                        try:
                            df_item = pd.read_csv(io.BytesIO(y_item_content), encoding='cp932')
                        except Exception:
                            df_item = pd.read_csv(io.BytesIO(y_item_content), encoding='utf-8')
                        df_item.columns = [str(c).strip() for c in df_item.columns]
                        df_item["対象月"] = extract_month_from_str(y_item_name)
                    
                    df_item = df_item.rename(columns={
                        "SKU": "商品コード",
                        "タイトル": "商品名",
                        "注文商品の売上額": "売上合計値（税込）",
                        "注文された商品点数": "注文点数"
                    }, errors='ignore')
                    df_item = df_item.loc[:, ~df_item.columns.duplicated()]
                    if "商品コード" in df_item.columns:
                        df_item["商品コード"] = df_item["商品コード"].fillna("不明").astype(str)
                    
                    for col in ["売上合計値（税込）", "注文点数"]:
                        if col in df_item.columns:
                            df_item[col] = df_item[col].astype(str).str.replace(',', '', regex=False).str.replace('円', '', regex=False).str.replace('¥', '', regex=False).str.replace('%', '', regex=False).str.strip()
                            df_item[col] = pd.to_numeric(df_item[col], errors='coerce').fillna(0.0)
                        else:
                            df_item[col] = 0.0
                    
                    df_item["店舗名"] = item_store_name
                    yahoo_item_dfs.append(df_item)
                else:
                    if y_item_name.lower().endswith('.xlsx'):
                        sheets_dict = pd.read_excel(io.BytesIO(y_item_content), sheet_name=None)
                        y_dfs = []
                        for sheet_name, df_sheet in sheets_dict.items():
                            df_sheet.columns = [str(c).strip() for c in df_sheet.columns]
                            m_val = extract_month_from_str(sheet_name)
                            if m_val == "不明":
                                m_val = extract_month_from_str(y_item_name)
                            df_sheet["対象月"] = m_val
                            y_dfs.append(df_sheet)
                        df_item = pd.concat(y_dfs, ignore_index=True)
                    else:
                        try:
                            df_item = pd.read_csv(io.BytesIO(y_item_content), encoding='cp932')
                        except Exception:
                            df_item = pd.read_csv(io.BytesIO(y_item_content), encoding='utf-8')
                        df_item.columns = [str(c).strip() for c in df_item.columns]
                        df_item["対象月"] = extract_month_from_str(y_item_name)
                    
                    df_item = df_item.rename(columns={
                        "注文点数合計": "注文点数"
                    }, errors='ignore')
                    
                    df_item = df_item.loc[:, ~df_item.columns.duplicated()]
                    if "商品コード" in df_item.columns:
                        df_item["商品コード"] = df_item["商品コード"].fillna("不明").astype(str)
                    
                    item_store_name = None
                    if "YS1" in y_item_name:
                        item_store_name = "Yahoo1号店"
                    elif "YS2" in y_item_name:
                        item_store_name = "Yahoo2号店"
                    else:
                        item_store_name = f"不明商品店舗_{y_item_name.replace('.xlsx', '').replace('.csv', '')}"
                    
                    df_item["店舗名"] = item_store_name
                    yahoo_item_dfs.append(df_item)
            except Exception as e:
                errors_list.append(f"❌ 商品別データの読み込みエラー ({y_item_name}): {e}")

    for up_name, up_content in (uploaded_files_data or []):
        try:
            store_name = get_store_name(up_name)
            sheets_dict = pd.read_excel(io.BytesIO(up_content), sheet_name=None)
            
            budget_dict = {}
            is_amazon = ("Amazon" in up_name or "UBUNBASE" in up_name)
            
            # 1. 予算シートを先行して読み込む
            if "予算" in sheets_dict:
                df_budget = sheets_dict["予算"]
                df_budget.columns = [str(c).strip() for c in df_budget.columns]
                for _, row in df_budget.iterrows():
                    month_name = str(row.iloc[0]).strip()
                    budget_amt = pd.to_numeric(row.iloc[1], errors='coerce')
                    if not pd.isna(budget_amt):
                        if is_amazon:
                            if month_name not in amazon_budget_loaded_months:
                                budget_dict[month_name] = budget_amt
                                amazon_budget_loaded_months.add(month_name)
                            else:
                                budget_dict[month_name] = 0.0
                        else:
                            budget_dict[month_name] = budget_amt
            
            for sheet_name, df in sheets_dict.items():
                if sheet_name == "予算":
                    continue
                
                df.columns = [str(c).strip() for c in df.columns]
                
                if is_amazon:
                    if 'まとめ' in sheet_name or sheet_name == "全体データ":
                        continue
                    if df.empty:
                        continue
                    
                    try:
                        def parse_amazon_date(val):
                            if pd.isna(val):
                                return pd.NaT
                            val_str = str(val).strip()
                            if val_str.replace('.', '', 1).isdigit() and len(val_str) < 8:
                                try:
                                    serial_num = float(val_str)
                                    return pd.to_datetime(serial_num, unit='D', origin='1899-12-30').normalize()
                                except Exception:
                                    pass
                            val_clean = val_str.split('(')[0].strip()
                            return pd.to_datetime(val_clean, errors='coerce').normalize()
                        
                        def clean_numeric_col(df_target, col_name, fallback_cols):
                            target_col = None
                            for col in df_target.columns:
                                if col == col_name or col in fallback_cols:
                                    target_col = col
                                    break
                            if target_col:
                                return pd.to_numeric(
                                    df_target[target_col].astype(str).str.replace(',', '', regex=False).str.replace('円', '', regex=False).str.replace('¥', '', regex=False).str.strip(),
                                    errors='coerce'
                                ).fillna(0.0)
                            else:
                                return pd.Series(0.0, index=df_target.index)
                        
                        clean_cols = [str(col).replace(' ', '').replace('\n', '').replace('\r', '') for col in df.columns]
                        
                        df_temp = df.copy()
                        df_temp.columns = clean_cols
                        
                        is_ubunbase = False
                        if "商品売上" in clean_cols or "TotalACOS" in clean_cols or "広告経由売上" in clean_cols:
                            is_ubunbase = True
                        
                        if is_ubunbase:
                            date_col = None
                            for col in df_temp.columns:
                                if col in ["日", "日付", "年月日", "Date"]:
                                    date_col = col
                                    break
                            
                            if date_col:
                                df_temp["日付"] = df_temp[date_col].apply(parse_amazon_date)
                                df_temp = df_temp.dropna(subset=["日付"])
                                
                                ad_cost = clean_numeric_col(df_temp, "広告費", ["広告費用", "費用", "Spend", "Cost"])
                                ad_sales = clean_numeric_col(df_temp, "広告経由売上", ["広告経由売上高", "広告売上", "経由売上", "Sales"])
                                ad_clicks = clean_numeric_col(df_temp, "クリック数", ["クリック", "Clicks"])
                                ad_orders = clean_numeric_col(df_temp, "コンバージョン数", ["コンバージョン", "注文数", "Conversions", "Orders"])
                                item_sales = clean_numeric_col(df_temp, "商品売上", ["売上", "商品売上高", "全体売上", "Revenue"])
                                
                                amazon_clean = pd.DataFrame()
                                amazon_clean["日付"] = df_temp["日付"]
                                amazon_clean["店舗名"] = store_name
                                amazon_clean["広告費"] = ad_cost
                                amazon_clean["広告売上"] = ad_sales
                                amazon_clean["クリック数"] = ad_clicks
                                amazon_clean["注文数"] = ad_orders
                                amazon_clean["商品売上"] = item_sales
                                amazon_extracted_dfs.append(amazon_clean)
                                
                                amz_sales_df = pd.DataFrame()
                                amz_sales_df["日付"] = df_temp["日付"]
                                amz_sales_df["店舗名"] = store_name
                                amz_sales_df["売上合計値"] = item_sales
                                amz_sales_df["前年売上"] = 0.0
                                amz_sales_df["前年比"] = 0.0
                                amz_sales_df["注文数合計"] = ad_orders
                                amz_sales_df["YSイベント"] = ""
                                amz_sales_df["ボーナスストア"] = ""
                                
                                budget_val = 0.0
                                if sheet_name in budget_dict:
                                    budget_val = budget_dict[sheet_name]
                                amz_sales_df["月予算"] = budget_val
                                
                                wdays = {0: '月', 1: '火', 2: '水', 3: '木', 4: '金', 5: '土', 6: '日'}
                                amz_sales_df['曜日'] = amz_sales_df['日付'].dt.dayofweek.map(wdays).fillna("")
                                
                                essential_cols = ['日付', '曜日', '店舗名', '売上合計値', '前年売上', '前年比', '注文数合計', 'YSイベント', 'ボーナスストア', '月予算']
                                all_data_list.append(amz_sales_df[essential_cols])
                        
                        else:
                            date_col = None
                            for col in df_temp.columns:
                                if col in ["日付", "日", "年月日", "Date"]:
                                    date_col = col
                                    break
                            
                            if date_col:
                                df_temp["日付"] = df_temp[date_col].apply(parse_amazon_date)
                                df_temp = df_temp.dropna(subset=["日付"])
                                
                                item_sales = clean_numeric_col(df_temp, "選択した日付の範囲(注文商品の売上額)", ["商品売上", "売上", "売上合計値"])
                                orders_total = clean_numeric_col(df_temp, "選択した日付の範囲(注文点数)", ["注文数合計", "発送済み商品数", "注文点数"])
                                last_year_sales = clean_numeric_col(df_temp, "前年同期(注文商品の売上額)", ["前年売上"])
                                
                                amazon_clean = pd.DataFrame()
                                amazon_clean["日付"] = df_temp["日付"]
                                amazon_clean["店舗名"] = store_name
                                amazon_clean["広告費"] = 0.0
                                amazon_clean["広告売上"] = 0.0
                                amazon_clean["クリック数"] = 0.0
                                amazon_clean["注文数"] = 0.0
                                amazon_clean["商品売上"] = item_sales
                                amazon_extracted_dfs.append(amazon_clean)
                                
                                amz_sales_df = pd.DataFrame()
                                amz_sales_df["日付"] = df_temp["日付"]
                                amz_sales_df["店舗名"] = store_name
                                amz_sales_df["売上合計値"] = item_sales
                                amz_sales_df["前年売上"] = last_year_sales
                                amz_sales_df["前年比"] = 0.0
                                amz_sales_df["注文数合計"] = orders_total
                                amz_sales_df["YSイベント"] = ""
                                amz_sales_df["ボーナスストア"] = ""
                                
                                budget_val = 0.0
                                if sheet_name in budget_dict:
                                    budget_val = budget_dict[sheet_name]
                                amz_sales_df["月予算"] = budget_val
                                
                                wdays = {0: '月', 1: '火', 2: '水', 3: '木', 4: '金', 5: '土', 6: '日'}
                                amz_sales_df['曜日'] = amz_sales_df['日付'].dt.dayofweek.map(wdays).fillna("")
                                
                                essential_cols = ['日付', '曜日', '店舗名', '売上合計値', '前年売上', '前年比', '注文数合計', 'YSイベント', 'ボーナスストア', '月予算']
                                all_data_list.append(amz_sales_df[essential_cols])
                    
                    except Exception as e:
                        errors_list.append(f"❌ Amazonデータ抽出エラー (シート: {sheet_name}): {e}")
                    continue
                
                else:
                    if 'まとめ' in sheet_name or sheet_name == "全体データ":
                        continue
                    if df.empty:
                        continue
                    
                    if "楽天" in up_name:
                        rakuten_rename = {
                            "すべて 売上": "売上合計値",
                            "すべて 売上件数": "注文数合計",
                            "売上": "売上合計値",
                            "すべて売上": "売上合計値",
                            "すべて売上件数": "注文数合計",
                            "イベント": "YSイベント",
                            "イベント2": "ボーナスストア"
                        }
                        df = df.rename(columns=rakuten_rename, errors='ignore')
                    else:
                        yahoo_rename = {"注文数": "注文数合計", "注文点数": "注文点数合計"}
                        df = df.rename(columns=yahoo_rename, errors='ignore')
                    
                    df = df.loc[:, ~df.columns.duplicated()]
                    
                    if '日付' in df.columns:
                        df['日付'] = df['日付'].apply(lambda x: str(x).split('(')[0].strip())
                        df['日付'] = pd.to_datetime(df['日付'], errors='coerce').dt.normalize()
                        
                        wdays = {0: '月', 1: '火', 2: '水', 3: '木', 4: '金', 5: '土', 6: '日'}
                        df['曜日'] = df['日付'].dt.dayofweek.map(wdays).fillna("")
                    
                    clean_cols = ['売上合計値', '前年売上', '前年比', '注文数合計', '月予算']
                    for col in clean_cols:
                        if col in df.columns:
                            df[col] = df[col].astype(str).str.replace(',', '', regex=False).str.replace('円', '', regex=False).str.replace('¥', '', regex=False).str.replace('%', '', regex=False)
                            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                    
                    if "売上合計値" not in df.columns or "日付" not in df.columns:
                        continue
                    
                    for col in ["YSイベント", "ボーナスストア"]:
                        if col in df.columns:
                            df[col] = df[col].fillna("").astype(str).replace("nan", "")
                    
                    budget_val = 0
                    if sheet_name in budget_dict:
                        budget_val = budget_dict[sheet_name]
                    elif "月予算" in df.columns:
                        budget_val = df["月予算"].iloc[0]
                    
                    df["店舗名"] = store_name
                    df["月予算"] = budget_val
                    
                    essential_cols = ['日付', '曜日', '店舗名', '売上合計値', '前年売上', '前年比', '注文数合計', 'YSイベント', 'ボーナスストア', '月予算']
                    available_cols = [c for c in essential_cols if c in df.columns]
                    
                    all_data_list.append(df[available_cols])

        except Exception as e:
            errors_list.append(f"❌ {up_name} 読み込みエラー: {e}")
            
    # Yahoo!ショッピング商品データの結合
    yahoo_item_df = None
    if yahoo_item_dfs:
        try:
            yahoo_item_df = pd.concat(yahoo_item_dfs, ignore_index=True)
            
            # Arrowの型エラーを避けるため商品コード列を文字列型に統一
            if "商品コード" in yahoo_item_df.columns:
                yahoo_item_df["商品コード"] = yahoo_item_df["商品コード"].fillna("不明").astype(str)
                
            for col in ["売上合計値（税込）", "注文点数"]:
                if col in yahoo_item_df.columns:
                    yahoo_item_df[col] = yahoo_item_df[col].astype(str).str.replace(',', '', regex=False).str.replace('円', '', regex=False).str.replace('¥', '', regex=False).str.replace('%', '', regex=False).str.strip()
                    yahoo_item_df[col] = pd.to_numeric(yahoo_item_df[col], errors='coerce').fillna(0.0)
                else:
                    yahoo_item_df[col] = 0.0
            
            if "商品コード" not in yahoo_item_df.columns:
                yahoo_item_df["商品コード"] = "不明"
            if "商品名" not in yahoo_item_df.columns:
                yahoo_item_df["商品名"] = "不明"
            if "対象月" not in yahoo_item_df.columns:
                yahoo_item_df["対象月"] = "不明"
            else:
                yahoo_item_df["対象月"] = yahoo_item_df["対象月"].fillna("不明")
        except Exception as e:
            errors_list.append(f"❌ Yahoo!商品データの結合エラー: {e}")

    # 抽出したAmazon広告データの結合とグループ化
    amazon_ad_df = None
    if amazon_extracted_dfs:
        try:
            amazon_ad_all = pd.concat(amazon_extracted_dfs, ignore_index=True)
            amazon_ad_df = amazon_ad_all.groupby(["日付", "店舗名"])[["広告費", "広告売上", "クリック数", "注文数", "商品売上"]].sum().reset_index()
        except Exception as e:
            errors_list.append(f"❌ Amazon広告データの集計エラー: {e}")

    combined_df = None
    if all_data_list:
        combined_df = pd.concat(all_data_list, ignore_index=True)

    return combined_df, yahoo_item_df, ad_df, yahoo_ad_df, amazon_ad_df, warnings_list, errors_list


# ファイルアップローダー
st.write("### 📂 データ取り込み")

with st.expander("📂 売上データの取り込み (必須)", expanded=True):
    uploaded_files = st.file_uploader(
        "Yahoo!または楽天のExcelファイルを選択してください", 
        type=["xlsx"],
        accept_multiple_files=True,
        help="複数のモールデータを一括統合します。高度なデータクレンジング機能を搭載。"
    )

with st.expander("🎯 広告データの取り込み (任意)", expanded=False):
    col_ad1, col_ad2 = st.columns(2)
    with col_ad1:
        ad_uploaded_file = st.file_uploader(
            "楽天RPP広告レポート (CSVまたはExcel) を選択してください (任意)", 
            type=["csv", "xlsx"],
            help="アップロードすると、店舗別KPIマトリクスに広告費、広告売上、ROASが表示されます。"
        )
    with col_ad2:
        yahoo_ad_uploaded_files = st.file_uploader(
            "Yahoo!アイテムマッチ広告レポート (CSV/Excel) を選択してください (任意)", 
            type=["csv", "xlsx"],
            accept_multiple_files=True,
            help="アップロードすると、店舗別KPIマトリクスにYahoo!の広告費、広告売上、ROASが表示されます。"
        )

with st.expander("📦 Yahoo! / 楽天 / Amazonの商品別データの取り込み (任意)", expanded=False):
    yahoo_item_uploaded_files = st.file_uploader(
        "Yahoo! / 楽天 / Amazonの商品別データ (CSV/Excel) を選択してください", 
        type=["csv", "xlsx"],
        accept_multiple_files=True,
        help="アップロードすると、商品横展開分析などに商品別データが反映されます。"
    )

yahoo_item_df = None

if uploaded_files or yahoo_item_uploaded_files:
    uploaded_files_data = [(f.name, f.getvalue()) for f in (uploaded_files or [])]
    ad_uploaded_file_data = (ad_uploaded_file.name, ad_uploaded_file.getvalue()) if ad_uploaded_file else None
    yahoo_ad_uploaded_files_data = [(f.name, f.getvalue()) for f in (yahoo_ad_uploaded_files or [])]
    yahoo_item_uploaded_files_data = [(f.name, f.getvalue()) for f in (yahoo_item_uploaded_files or [])]

    with st.spinner("データを統合・計算中..."): 
        combined_df, yahoo_item_df, ad_df, yahoo_ad_df, amazon_ad_df, warnings_list, errors_list = load_and_process_data(
            uploaded_files_data, ad_uploaded_file_data, yahoo_ad_uploaded_files_data, yahoo_item_uploaded_files_data
        )
    
    # 警告とエラーをメイン処理側（関数の外）で表示
    for err in errors_list:
        st.error(err)
    for warn in warnings_list:
        st.warning(warn)
    
    if combined_df is not None and not combined_df.empty:
        
        # RPP広告データの結合（KPIマトリクスで直接集計するため、マージ処理は不要になりました）
        
        # 日付データの最終処理と累計計算
        if "日付" in combined_df.columns:
            combined_df = combined_df.dropna(subset=["日付"])
            combined_df["年月"] = combined_df["日付"].dt.strftime('%Y年%m月')
            combined_df["純粋な日付"] = combined_df["日付"].dt.date
            
            # 週ラベルの作成（月曜始まりの週範囲：YYYY/MM/DD 〜 MM/DD）
            monday = combined_df["日付"] - pd.to_timedelta(combined_df["日付"].dt.dayofweek, unit='D')
            sunday = monday + pd.to_timedelta(6, unit='D')
            combined_df["週ラベル"] = monday.dt.strftime('%Y/%m/%d') + " 〜 " + sunday.dt.strftime('%m/%d')
            
            combined_df = combined_df.sort_values(["店舗名", "純粋な日付"])
            combined_df["売上累計"] = combined_df.groupby(["店舗名", "年月"])["売上合計値"].cumsum()
            combined_df["前年比"] = operator.mul(combined_df["売上合計値"] / combined_df["前年売上"], 100).replace([float('inf'), -float('inf')], 0).fillna(0)
            combined_df["予算比"] = operator.mul(combined_df["売上累計"] / combined_df["月予算"], 100).replace([float('inf'), -float('inf')], 0).fillna(0)
            
            # サイドバー設定
            with st.sidebar:
                st.header("⚙️ 表示設定")
                period_type = st.radio("集計期間", ["月次", "週次"], index=0)
                
                selected_month = None
                selected_week = None
                
                if period_type == "月次":
                    available_months = sorted(combined_df["年月"].unique(), reverse=True)
                    selected_month = st.selectbox("表示月を選択", available_months, index=0) if available_months else None
                else:
                    available_weeks = sorted(combined_df["週ラベル"].unique(), reverse=True)
                    selected_week = st.selectbox("表示週を選択", available_weeks, index=0) if available_weeks else None
                
                display_mode = st.selectbox("表示モード", ["全店合計", "個別店舗"])
                
                st.write("---")
                st.header("📷 画面キャプチャ")
                import streamlit.components.v1 as components
                capture_btn_html = """
                <script src="https://cdnjs.cloudflare.com/ajax/libs/dom-to-image-more/3.1.6/dom-to-image-more.min.js"></script>
                <button onclick="downloadImage()" style="width: 100%; padding: 0.6rem; background-color: #00f2fe; color: #0e1117; border: none; border-radius: 0.3rem; font-family: sans-serif; font-size: 14px; font-weight: 700; cursor: pointer; box-shadow: 0 0 10px rgba(0, 242, 254, 0.4); transition: 0.2s;">
                    📷 画面全体をキャプチャ
                </button>
                <script>
                function downloadImage() {
                    const btn = document.querySelector('button');
                    const originalText = btn.innerText;
                    btn.innerText = "⏳ キャプチャ中...";
                    const parentDoc = window.parent.document;
                    let scrollArea = parentDoc.querySelector('.main') || parentDoc.querySelector('[data-testid="stAppViewBlockContainer"]') || parentDoc.documentElement;
                    let target = parentDoc.querySelector('[data-testid="stMainBlockContainer"]') || parentDoc.querySelector('[data-testid="stAppViewBlockContainer"]') || parentDoc.querySelector('.main') || parentDoc.body;
                    
                    if(target && scrollArea) {
                        const currentBgColor = window.getComputedStyle(parentDoc.body).backgroundColor;
                        setTimeout(() => {
                            window.domtoimage.toPng(target, {
                                bgcolor: currentBgColor || "#0f172a",
                                width: target.scrollWidth,
                                height: target.scrollHeight,
                                cacheBust: true
                            }).then(function (dataUrl) {
                                const link = parentDoc.createElement('a');
                                link.download = 'Cerberus_Dashboard.png';
                                link.href = dataUrl;
                                link.click();
                                btn.innerText = "✅ 保存完了！";
                                setTimeout(() => { btn.innerText = originalText; }, 3000);
                            }).catch(function (error) {
                                console.error("Capture Error:", error);
                                btn.innerText = "❌ エラー";
                                setTimeout(() => { btn.innerText = originalText; }, 3000);
                            });
                        }, 1000);
                    } else {
                        alert("メイン画面が見つかりませんでした。");
                        btn.innerText = originalText;
                    }
                }
                </script>
                """
                components.html(capture_btn_html, height=80)
            
            # フィルタリングと先週基準日の特定
            last_week_start = None
            last_week_end = None
            
            if period_type == "月次":
                filtered_df = combined_df[combined_df["年月"] == selected_month] if selected_month else combined_df
            else:
                filtered_df = combined_df[combined_df["週ラベル"] == selected_week] if selected_week else combined_df
                if selected_week:
                    start_str = selected_week.split(" 〜 ")[0]
                    this_week_start = pd.to_datetime(start_str)
                    last_week_start = this_week_start - pd.to_timedelta(7, unit='D')
                    last_week_end = this_week_start - pd.to_timedelta(1, unit='D')

            # Plotlyグラフ描画
            def draw_plotly_chart(df, title):
                graph_df = df.copy()
                graph_df["日付ラベル"] = pd.to_datetime(graph_df["純粋な日付"]).dt.strftime('%Y-%m-%d')
                
                fig = go.Figure()
                fig.add_trace(go.Scatter(
                    x=graph_df["日付ラベル"], y=graph_df["前年売上"],
                    mode='lines', name='前年売上',
                    line=dict(color='#FFD700', width=2, dash='dash'),
                    hoverinfo='skip'
                ))
                fig.add_trace(go.Scatter(
                    x=graph_df["日付ラベル"], y=graph_df["売上合計値"],
                    mode='lines+markers', name='今年売上',
                    line=dict(color='#00f2fe', width=3),
                    marker=dict(size=6),
                    customdata=graph_df[["前年売上", "YSイベント", "ボーナスストア"]] if all(c in graph_df.columns for c in ["前年売上", "YSイベント", "ボーナスストア"]) else None,
                    hovertemplate="<b>%{x}</b><br>" +
                                  "今年売上: ¥%{y:,.0f}<br>" +
                                  "前年売上: ¥%{customdata[0]:,.0f}<br>" +
                                  "イベント: %{customdata[1]}<br>" +
                                  "ボーナス: %{customdata[2]}<extra></extra>"
                ))
                if "YSイベント" in graph_df.columns and "ボーナスストア" in graph_df.columns:
                    event_df = graph_df[(graph_df["YSイベント"] != "") | (graph_df["ボーナスストア"] != "")]
                    if not event_df.empty:
                        fig.add_trace(go.Scatter(
                            x=event_df["日付ラベル"], y=event_df["売上合計値"],
                            mode='markers', name='イベント発生',
                            marker=dict(size=12, color='#ff00ff', symbol='star'),
                            hoverinfo='skip'
                        ))
                
                fig.update_layout(
                    title=title, template="plotly_dark", hovermode="x unified",
                    xaxis=dict(title="日付", tickangle=45),
                    yaxis=dict(title="売上金額 (円)", tickformat=",d"),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                    height=500, margin=dict(l=20, r=20, t=60, b=20)
                )
                st.plotly_chart(fig, use_container_width=True)

            # 表示用フォーマット
            def format_dataframe(df):
                display_cols = ['日付表示', '曜日', '店舗名', '売上合計値', '前年売上', '前年比', '注文数合計', '月予算', '予算比']
                df_prep = df.copy()
                df_prep['日付表示'] = pd.to_datetime(df_prep['純粋な日付']).dt.strftime('%Y-%m-%d')
                actual_cols = [c for c in display_cols if c in df_prep.columns]
                df_final = df_prep[actual_cols].fillna(0)
                
                return df_final.style.format({
                    "売上合計値": "{:,.0f}", "前年売上": "{:,.0f}", "注文数合計": "{:,.0f}",
                    "月予算": "{:,.0f}", "前年比": "{:.1f}%", "予算比": "{:.1f}%"
                })

            # 🔮 月末着地予想・目標逆算の共通描画関数
            def show_monthly_forecast_and_target_backcast(target_month, selected_stores_list=None):
                st.write("### 🔮 月末着地予想・目標逆算")
                
                # 対象月のデータ全体を安全に抽出
                month_full_df = combined_df[combined_df["年月"] == target_month]
                if month_full_df.empty:
                    st.info("ℹ️ 着地予想データ不足（対象月のデータが存在しません）")
                    return
                
                # 個別店舗モードでの店舗フィルタリング
                if selected_stores_list:
                    month_store_df = month_full_df[month_full_df["店舗名"].isin(selected_stores_list)]
                else:
                    month_store_df = month_full_df
                
                if month_store_df.empty:
                    st.info("ℹ️ 着地予想データ不足（選択店舗の対象月データが存在しません）")
                    return
                
                # 重複のない日付単位でグループ化して集計
                sum_cols = ["売上合計値", "月予算"]
                agg_dict = {c: "sum" for c in sum_cols if c in month_store_df.columns}
                daily_month_df = month_store_df.groupby("純粋な日付").agg(agg_dict).reset_index()
                
                # 1. 経過日数の取得：売上が0より大きい有効な最新日付の「日」を取得
                valid_df = daily_month_df[daily_month_df["売上合計値"] > 0]
                if not valid_df.empty:
                    elapsed_days = pd.to_datetime(valid_df["純粋な日付"].max()).day
                else:
                    elapsed_days = 1
                
                # 2. その月の総日数の取得：calendarモジュールから正確な月末日を算出
                import calendar
                try:
                    year = int(target_month.split("年")[0].strip())
                    month = int(target_month.split("年")[1].split("月")[0].strip())
                    _, total_days = calendar.monthrange(year, month)
                except Exception:
                    total_days = 30
                
                # 3. 残り日数
                remaining_days = total_days - elapsed_days
                if remaining_days <= 0:
                    remaining_days = 1
                
                # 4. 指標の計算
                current_sales = daily_month_df["売上合計値"].sum()
                monthly_budget = daily_month_df["月予算"].max() if not daily_month_df.empty else 0
                
                # 月末着地予想：(現在の総売上 / 経過日数) x その月の総日数
                daily_avg = current_sales / elapsed_days
                forecast_sales = operator.mul(daily_avg, total_days)
                
                sales_diff = forecast_sales - monthly_budget
                forecast_achieve_rate = operator.mul(forecast_sales / monthly_budget, 100) if monthly_budget > 0 else 0
                
                required_sales = monthly_budget - current_sales
                required_daily_sales = required_sales / remaining_days if remaining_days > 0 else 0
                if required_daily_sales < 0:
                    required_daily_sales = 0
                
                # 5. UIの描画
                cols = st.columns(3)
                cols[0].metric(
                    "月末着地予想", 
                    f"¥{int(forecast_sales):,}", 
                    delta=f"予算比 {int(sales_diff):+,d}円" if monthly_budget > 0 else None
                )
                cols[1].metric("着地達成率", f"{forecast_achieve_rate:.1f}%")
                cols[2].metric("達成に必要な今日からの日販", f"¥{int(required_daily_sales):,}")

            # 🚀 施策（イベント）リフト分析の共通描画関数 (ユーザー指定コードの完全置換)
            def show_event_lift_analysis(target_df):
                st.write("### 🚀 施策（イベント）リフト分析")
                
                # 1. 判定用の安全なクレンジング（NaNや文字の"nan"を強制的に空文字にする）
                ys_ev = target_df.get('YSイベント', pd.Series(dtype=str)).fillna('').astype(str).str.strip().replace(['nan', 'None', 'NaN'], '')
                bonus_ev = target_df.get('ボーナスストア', pd.Series(dtype=str)).fillna('').astype(str).str.strip().replace(['nan', 'None', 'NaN'], '')
                
                # 2. イベント日と通常日の判定
                is_event = (ys_ev != '') | (bonus_ev != '')
                event_days_df = target_df[is_event]
                normal_days_df = target_df[~is_event]
                
                # 3. リフト値の計算と表示
                if len(event_days_df) == 0:
                    st.info("比較データ不足（この期間内にイベント開催日がありません）")
                elif len(normal_days_df) == 0:
                    st.info("比較データ不足（すべての日がイベント対象のため、通常日との比較ができません）")
                else:
                    event_avg = event_days_df['売上合計値'].mean()
                    normal_avg = normal_days_df['売上合計値'].mean()
                    lift_ratio = event_avg / normal_avg if normal_avg > 0 else 0
                    
                    cols = st.columns(3)
                    cols[0].metric("通常日 1日平均売上", f"¥{int(normal_avg):,}")
                    cols[1].metric("イベント日 1日平均売上", f"¥{int(event_avg):,}")
                    cols[2].metric("売上リフト効果", f"{lift_ratio:.1f} 倍")

            if display_mode == "全店合計":
                title_period = selected_month if period_type == "月次" else selected_week
                
                # --- 事前集計: 売上パフォーマンス ---
                sum_cols = ["売上合計値", "前年売上", "売上累計", "月予算", "注文数合計"]
                agg_dict = {c: "sum" for c in sum_cols if c in filtered_df.columns}
                if "YSイベント" in filtered_df.columns:
                    agg_dict["YSイベント"] = lambda x: ", ".join([str(v) for v in x.unique() if pd.notna(v) and str(v).strip() not in ["", "nan", "None"]])
                if "ボーナスストア" in filtered_df.columns:
                    agg_dict["ボーナスストア"] = lambda x: ", ".join([str(v) for v in x.unique() if pd.notna(v) and str(v).strip() not in ["", "nan", "None"]])
                if "曜日" in filtered_df.columns:
                    agg_dict["曜日"] = "first"
                
                summary_df = filtered_df.groupby("純粋な日付").agg(agg_dict).reset_index()
                summary_df["月予算"] = summary_df["月予算"].max()
                summary_df["前年比"] = operator.mul(summary_df["売上合計値"] / summary_df["前年売上"], 100).fillna(0)
                summary_df["予算比"] = operator.mul(summary_df["売上累計"] / summary_df["月予算"], 100).replace([float('inf'), -float('inf')], 0).fillna(0)
                
                this_week_sales = summary_df['売上合計値'].sum() if not summary_df.empty else 0
                this_week_orders = summary_df['注文数合計'].sum() if not summary_df.empty else 0
                
                sales_delta = None
                orders_delta = None
                
                if period_type == "週次" and last_week_start is not None:
                    # 全店舗の先週のデータを取得
                    last_week_df = combined_df[(combined_df["日付"] >= last_week_start) & (combined_df["日付"] <= last_week_end)]
                    last_week_sales = last_week_df["売上合計値"].sum() if not last_week_df.empty else 0
                    last_week_orders = last_week_df["注文数合計"].sum() if not last_week_df.empty else 0
                    
                    if last_week_sales > 0:
                        sales_diff_pct = operator.mul((this_week_sales - last_week_sales) / last_week_sales, 100)
                        sales_delta = f"先週比 {sales_diff_pct:+.1f}% (先週: ¥{int(last_week_sales):,})"
                    else:
                        sales_delta = "先週比 N/A (先週実績なし)"
                        
                    if last_week_orders > 0:
                        orders_diff_pct = operator.mul((this_week_orders - last_week_orders) / last_week_orders, 100)
                        orders_delta = f"先週比 {orders_diff_pct:+.1f}% (先週: {int(last_week_orders):,}件)"
                    else:
                        orders_delta = "先週比 N/A (先週実績なし)"
                
                # --- 事前集計: 店舗別 KPIマトリクス ---
                import numpy as np
                filter_dates = pd.to_datetime(filtered_df["日付"], errors='coerce').dt.normalize()
                min_date = filter_dates.min()
                max_date = filter_dates.max()
                
                ad_cost_total = np.nan
                ad_sales_total = np.nan
                ad_roas = np.nan
                
                if ad_df is not None and not ad_df.empty and pd.notna(min_date) and pd.notna(max_date):
                    ad_df_temp = ad_df.copy()
                    ad_df_temp["日付_normalized"] = pd.to_datetime(ad_df_temp["日付"], errors='coerce').dt.normalize()
                    ad_filtered = ad_df_temp[(ad_df_temp["日付_normalized"] >= min_date) & (ad_df_temp["日付_normalized"] <= max_date)]
                    if not ad_filtered.empty:
                        ad_cost_total = ad_filtered["広告費"].sum()
                        ad_sales_total = ad_filtered["広告売上"].sum()
                        ad_roas = operator.mul(ad_sales_total / ad_cost_total, 100) if ad_cost_total > 0 else 0.0
                
                yahoo_ad_summary = {}
                if yahoo_ad_df is not None and not yahoo_ad_df.empty and pd.notna(min_date) and pd.notna(max_date):
                    yahoo_ad_df_temp = yahoo_ad_df.copy()
                    yahoo_ad_df_temp["日付_normalized"] = pd.to_datetime(yahoo_ad_df_temp["日付"], errors='coerce').dt.normalize()
                    y_ad_filtered = yahoo_ad_df_temp[(yahoo_ad_df_temp["日付_normalized"] >= min_date) & (yahoo_ad_df_temp["日付_normalized"] <= max_date)]
                    if not y_ad_filtered.empty:
                        y_grp = y_ad_filtered.groupby("店舗")
                        for store_name, group in y_grp:
                            cost = group["広告費"].sum()
                            sales = group["広告売上"].sum()
                            roas = operator.mul(sales / cost, 100) if cost > 0 else 0.0
                            yahoo_ad_summary[store_name] = {"cost": cost, "sales": sales, "roas": roas}
                
                amazon_ad_cost = np.nan
                amazon_ad_sales = np.nan
                amazon_ad_roas = np.nan
                if amazon_ad_df is not None and not amazon_ad_df.empty and pd.notna(min_date) and pd.notna(max_date):
                    amazon_ad_df_temp = amazon_ad_df.copy()
                    amazon_ad_df_temp["日付_normalized"] = pd.to_datetime(amazon_ad_df_temp["日付"], errors='coerce').dt.normalize()
                    amazon_filtered = amazon_ad_df_temp[(amazon_ad_df_temp["日付_normalized"] >= min_date) & (amazon_ad_df_temp["日付_normalized"] <= max_date)]
                    if not amazon_filtered.empty:
                        amazon_ad_cost = amazon_filtered["広告費"].sum()
                        amazon_ad_sales = amazon_filtered["広告売上"].sum()
                        amazon_ad_roas = operator.mul(amazon_ad_sales / amazon_ad_cost, 100) if amazon_ad_cost > 0 else 0.0
                
                kpi_base_df = filtered_df.copy()
                store_grp = kpi_base_df.groupby("店舗名")
                
                kpi_df = pd.DataFrame()
                kpi_df["売上"] = store_grp["売上合計値"].sum()
                kpi_df["前年売上"] = store_grp["前年売上"].sum()
                kpi_df["前年比"] = operator.mul(kpi_df["売上"] / kpi_df["前年売上"], 100).replace([float('inf'), -float('inf')], 0).fillna(0)
                kpi_df["注文数"] = store_grp["注文数合計"].sum()
                kpi_df["客単価"] = (kpi_df["売上"] / kpi_df["注文数"]).replace([float('inf'), -float('inf')], 0).fillna(0)
                
                daily_budget = kpi_base_df.groupby(["店舗名", "純粋な日付"])["月予算"].sum()
                kpi_df["月予算"] = daily_budget.groupby(level="店舗名").max().fillna(0)
                kpi_df["予算進捗"] = operator.mul(kpi_df["売上"] / kpi_df["月予算"], 100).replace([float('inf'), -float('inf')], 0).fillna(0)
                
                kpi_df["広告費"] = np.nan
                kpi_df["広告売上"] = np.nan
                kpi_df["ROAS"] = np.nan
                
                if "楽天" in kpi_df.index:
                    if pd.notna(ad_cost_total):
                         kpi_df.loc["楽天", "広告費"] = ad_cost_total
                         kpi_df.loc["楽天", "広告売上"] = ad_sales_total
                         kpi_df.loc["楽天", "ROAS"] = ad_roas
                    else:
                         kpi_df.loc["楽天", "広告費"] = 0.0
                         kpi_df.loc["楽天", "広告売上"] = 0.0
                         kpi_df.loc["楽天", "ROAS"] = 0.0
                
                for store_name in kpi_df.index:
                    if store_name.startswith("Yahoo"):
                        if store_name in yahoo_ad_summary:
                            kpi_df.loc[store_name, "広告費"] = yahoo_ad_summary[store_name]["cost"]
                            kpi_df.loc[store_name, "広告売上"] = yahoo_ad_summary[store_name]["sales"]
                            kpi_df.loc[store_name, "ROAS"] = yahoo_ad_summary[store_name]["roas"]
                        else:
                            kpi_df.loc[store_name, "広告費"] = 0.0
                            kpi_df.loc[store_name, "広告売上"] = 0.0
                            kpi_df.loc[store_name, "ROAS"] = 0.0
                    elif store_name in ["Amazon1号店", "Amazon2号店", "Amazon合計", "Amazon"]:
                        store_ad_cost = np.nan
                        store_ad_sales = np.nan
                        store_ad_roas = np.nan
                        if amazon_ad_df is not None and not amazon_ad_df.empty:
                            amazon_ad_df_temp = amazon_ad_df.copy()
                            amazon_ad_df_temp["日付_normalized"] = pd.to_datetime(amazon_ad_df_temp["日付"], errors='coerce').dt.normalize()
                            amz_filtered = amazon_ad_df_temp[(amazon_ad_df_temp["日付_normalized"] >= min_date) & (amazon_ad_df_temp["日付_normalized"] <= max_date) & (amazon_ad_df_temp["店舗名"] == store_name)]
                            if not amz_filtered.empty:
                                store_ad_cost = amz_filtered["広告費"].sum()
                                store_ad_sales = amz_filtered["広告売上"].sum()
                                store_ad_roas = operator.mul(store_ad_sales / store_ad_cost, 100) if store_ad_cost > 0 else 0.0
                        
                        if pd.notna(store_ad_cost):
                            kpi_df.loc[store_name, "広告費"] = store_ad_cost
                            kpi_df.loc[store_name, "広告売上"] = store_ad_sales
                            kpi_df.loc[store_name, "ROAS"] = store_ad_roas
                        else:
                            kpi_df.loc[store_name, "広告費"] = 0.0
                            kpi_df.loc[store_name, "広告売上"] = 0.0
                            kpi_df.loc[store_name, "ROAS"] = 0.0
                
                # Amazonを含む店舗を抽出
                amazon_stores = [idx for idx in kpi_df.index if "Amazon" in str(idx)]
                if len(amazon_stores) >= 2:
                    # 単純合算する数値を計算
                    sum_sales = kpi_df.loc[amazon_stores, "売上"].sum()
                    sum_prev_sales = kpi_df.loc[amazon_stores, "前年売上"].sum()
                    sum_orders = kpi_df.loc[amazon_stores, "注文数"].sum()
                    sum_budget = kpi_df.loc[amazon_stores, "月予算"].sum()
                    sum_ad_cost = kpi_df.loc[amazon_stores, "広告費"].sum()
                    sum_ad_sales = kpi_df.loc[amazon_stores, "広告売上"].sum()
                    
                    # 各指標の再計算（アスタリスク記号は一切禁止）
                    calc_prev_ratio = 0.0
                    if sum_prev_sales > 0:
                        calc_prev_ratio = operator.mul(sum_sales / sum_prev_sales, 100)
                    
                    calc_avg_order = 0.0
                    if sum_orders > 0:
                        calc_avg_order = sum_sales / sum_orders
                        
                    calc_budget_progress = 0.0
                    if sum_budget > 0:
                        calc_budget_progress = operator.mul(sum_sales / sum_budget, 100)
                        
                    calc_roas = 0.0
                    if sum_ad_cost > 0:
                        calc_roas = operator.mul(sum_ad_sales / sum_ad_cost, 100)
                        
                    # Amazon合計行を追加
                    kpi_df.loc["Amazon合計"] = [
                        sum_sales,
                        sum_prev_sales,
                        calc_prev_ratio,
                        sum_orders,
                        calc_avg_order,
                        sum_budget,
                        calc_budget_progress,
                        sum_ad_cost,
                        sum_ad_sales,
                        calc_roas
                    ]
                
                # 個別のAmazon店舗の月予算と予算進捗を無効化（np.nanを代入）
                for amz_store in amazon_stores:
                    if amz_store in kpi_df.index:
                        kpi_df.loc[amz_store, "月予算"] = np.nan
                        kpi_df.loc[amz_store, "予算進捗"] = np.nan

                kpi_cols = ["売上", "前年売上", "前年比", "注文数", "客単価", "月予算", "予算進捗", "広告費", "広告売上", "ROAS"]
                kpi_df = kpi_df[kpi_cols]
                
                # --- タブUIの実装 ---
                tab_sales, tab_ads, tab_items = st.tabs(["📈 売上分析", "🎯 広告分析", "📦 商品横展開分析"])

                
                with tab_sales:
                    st.write(f"### 📊 全店合計パフォーマンス ({title_period})")
                    
                    cols = st.columns(3)
                    cols[0].metric("総売上", f"¥{int(this_week_sales):,}", delta=sales_delta)
                    cols[1].metric("予算進捗（最新）", f"{summary_df['予算比'].iloc[-1] if not summary_df.empty else 0:.1f}%")
                    cols[2].metric("総注文数", f"{int(this_week_orders):,} 件", delta=orders_delta)
                    
                    # 月末着地予想・目標逆算
                    st.markdown("---")
                    target_month = selected_month if period_type == "月次" else pd.to_datetime(selected_week.split(" 〜 ")[0]).strftime('%Y年%m月')
                    show_monthly_forecast_and_target_backcast(target_month)
                    
                    # 施策リフト分析
                    st.markdown("---")
                    show_event_lift_analysis(summary_df)
                    
                    st.markdown("---")
                    draw_plotly_chart(summary_df, "全モール合計 売上推移")
                    
                    # モール別売上シェアの円グラフ
                    if filtered_df is not None and not filtered_df.empty and "店舗名" in filtered_df.columns:
                        store_share = filtered_df.groupby("店舗名")["売上合計値"].sum().reset_index()
                        if store_share["売上合計値"].sum() > 0:
                            fig_pie = go.Figure()
                            fig_pie.add_trace(go.Pie(
                                labels=store_share["店舗名"],
                                values=store_share["売上合計値"],
                                texttemplate="%{percent:.1%} / ¥%{value:,.0f}",
                                textposition="inside",
                                hovertemplate="<b>%{label}</b><br>売上金額: ¥%{value:,.0f}<br>構成比: %{percent:.1%}<extra></extra>"
                            ))
                            fig_pie.update_layout(
                                title="モール別 売上シェア",
                                template="plotly_dark",
                                height=450,
                                margin=dict(l=20, r=20, t=60, b=20),
                                legend=dict(orientation="h", yanchor="bottom", y=-0.1, xanchor="center", x=0.5)
                            )
                            st.plotly_chart(fig_pie, use_container_width=True)
                    
                    st.dataframe(format_dataframe(summary_df), use_container_width=True)
                    
                    # 店舗別 KPIマトリクス
                    st.markdown("---")
                    st.write("### 📊 店舗別 KPIマトリクス")
                    st.dataframe(
                        kpi_df.style.format({
                            "売上": "¥{:,.0f}",
                            "前年売上": "¥{:,.0f}",
                            "前年比": "{:.1f}%",
                            "注文数": "{:,.0f}",
                            "客単価": "¥{:,.0f}",
                            "月予算": "¥{:,.0f}",
                            "予算進捗": "{:.1f}%",
                            "広告費": "¥{:,.0f}",
                            "広告売上": "¥{:,.0f}",
                            "ROAS": "{:.1f}%"
                        }, na_rep="-"),
                        use_container_width=True
                    )
                
                with tab_ads:
                    # --- 📊 広告費と売上の相関分析 (楽天) ---
                    st.write("### 📊 広告費と売上の相関分析 (楽天)")
                    if ad_df is not None and not ad_df.empty:
                        rakuten_sales_daily = filtered_df[filtered_df["店舗名"] == "楽天"].groupby("日付")["売上合計値"].sum().reset_index()
                        
                        if not rakuten_sales_daily.empty:
                            ad_df_temp = ad_df.copy()
                            ad_df_temp["日付"] = pd.to_datetime(ad_df_temp["日付"]).dt.normalize()
                            rakuten_sales_daily["日付"] = pd.to_datetime(rakuten_sales_daily["日付"]).dt.normalize()
                            
                            scatter_data = pd.merge(rakuten_sales_daily, ad_df_temp[["日付", "広告費", "広告売上", "クリック数", "注文数"]], on="日付", how="left")
                            scatter_data["広告費"] = pd.to_numeric(scatter_data["広告費"], errors='coerce').fillna(0)
                            scatter_data["広告売上"] = pd.to_numeric(scatter_data["広告売上"], errors='coerce').fillna(0)
                            scatter_data["クリック数"] = pd.to_numeric(scatter_data["クリック数"], errors='coerce').fillna(0)
                            scatter_data["注文数"] = pd.to_numeric(scatter_data["注文数"], errors='coerce').fillna(0)
                            
                            if not scatter_data.empty:
                                st.markdown("""
                                    <style>
                                    [data-testid="stMetricLabel"] {
                                        width: 100% !important;
                                        min-width: 120px !important;
                                    }
                                    [data-testid="stMetricLabel"] > div,
                                    [data-testid="stMetricLabel"] p,
                                    [data-testid="stMetricLabel"] span,
                                    [data-testid="stMetricLabel"] label {
                                        white-space: normal !important;
                                        overflow: visible !important;
                                        text-overflow: clip !important;
                                        display: inline-block !important;
                                        min-width: 100px !important;
                                    }
                                    </style>
                                    """, unsafe_allow_html=True)
                                st.subheader("📊 広告パフォーマンスサマリー")
                                r_total_cost = scatter_data["広告費"].sum()
                                r_total_sales = scatter_data["広告売上"].sum()
                                r_total_clicks = scatter_data["クリック数"].sum()
                                r_total_orders = scatter_data["注文数"].sum()
                                
                                r_roas = operator.mul(r_total_sales / r_total_cost, 100) if r_total_cost > 0 else 0.0
                                r_cpc = (r_total_cost / r_total_clicks) if r_total_clicks > 0 else 0.0
                                r_cvr = operator.mul(r_total_orders / r_total_clicks, 100) if r_total_clicks > 0 else 0.0
                                r_cpa = (r_total_cost / r_total_orders) if r_total_orders > 0 else 0.0
                                
                                r_cols_top = st.columns(3)
                                r_cols_top[0].metric("広告費", f"¥{int(r_total_cost):,}")
                                r_cols_top[1].metric("広告売上", f"¥{int(r_total_sales):,}")
                                r_cols_top[2].metric("ROAS", f"{r_roas:.1f}%")
                                
                                r_cols_bottom = st.columns(3)
                                r_cols_bottom[0].metric("CVR", f"{r_cvr:.1f}%")
                                r_cols_bottom[1].metric("CPC", f"¥{r_cpc:.1f}")
                                r_cols_bottom[2].metric("CPA", f"¥{int(r_cpa):,}")
                                
                                import numpy as np
                                x = scatter_data["広告費"].values
                                y = scatter_data["売上合計値"].values
                                
                                fig_scatter = go.Figure()
                                fig_scatter.add_trace(go.Scatter(
                                    x=x, y=y,
                                    mode='markers',
                                    name='日別実績',
                                    marker=dict(size=10, color='#00f2fe', opacity=0.8, line=dict(width=1, color='white')),
                                    text=scatter_data["日付"].dt.strftime('%Y-%m-%d'),
                                    hovertemplate="<b>日付: %{text}</b><br>広告費: ¥%{x:,.0f}<br>総売上: ¥%{y:,.0f}<extra></extra>"
                                ))
                                
                                if len(x) > 1 and np.var(x) > 0:
                                    try:
                                        slope, intercept = np.polyfit(x, y, 1)
                                        x_line = np.array([np.min(x), np.max(x)])
                                        y_line = operator.mul(slope, x_line) + intercept
                                        fig_scatter.add_trace(go.Scatter(
                                            x=x_line, y=y_line,
                                            mode='lines',
                                            name='傾向線 (回帰直線)',
                                            line=dict(color='#ff00ff', width=2, dash='dash'),
                                            hovertemplate="傾向線<extra></extra>"
                                        ))
                                    except Exception:
                                        pass
                                
                                fig_scatter.update_layout(
                                    title="広告費と総売上の相関分布",
                                    template="plotly_dark",
                                    xaxis=dict(title="広告費 (円)", tickformat=",d"),
                                    yaxis=dict(title="総売上 (円)", tickformat=",d"),
                                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                                    height=450, margin=dict(l=20, r=20, t=60, b=20)
                                )
                                st.plotly_chart(fig_scatter, use_container_width=True)
                                
                                correlation = scatter_data["広告費"].corr(scatter_data["売上合計値"])
                                if pd.notna(correlation):
                                    st.caption(f"💡 広告費と総売上の相関係数: {correlation:.2f} （1.0に近いほど強い正の相関があります。目安: 0.7以上で強い相関、0.4〜0.7で中程度の相関）")
                                    st.caption("※紫色の点線は『傾向線（回帰直線）』です。この線の傾きが右肩上がりであるほど、広告費の増減が売上に直接影響を与えている（広告が効果的に機能している）ことを示します。逆に線が横ばいや右下がりの場合は、広告費と売上の連動性が低く、運用の見直しが必要なサインとなります。")
                                    
                                    st.subheader("💡 システムからの戦術提案")
                                    current_roas = ad_roas if (ad_roas is not None and pd.notna(ad_roas)) else 0.0
                                    current_corr = correlation
                                    
                                    if current_corr >= 0.4 and current_roas >= 300.0:
                                        st.success("【投資推奨】広告と売上の連動性が高く、費用対効果も良好です。現在の着地予想を上回るために、日次予算の引き上げを検討してください。")
                                    elif current_corr < 0.2 and current_roas >= 300.0:
                                        st.warning("【予算削減・再配分推奨】広告のROASは高いですが、広告費と全体売上が連動していません。自然検索で売れるはずの注文を広告で獲得している（カニバリ）可能性があります。入札単価を下げて無駄なコストを削り、利益を確保してください。")
                                    elif current_corr < 0.2 and current_roas < 300.0:
                                        st.warning("【設定見直し推奨】広告費の増加が全体の売上増加に繋がっていません。無駄なクリックが発生している可能性があるため、RPPの除外キーワード設定や、費用対効果の高い他モールへの予算シフトを検討してください。")
                                    elif current_corr >= 0.4 and current_roas < 300.0:
                                        st.info("【CVR改善推奨】広告による集客効果は出ていますが、ROASが基準を下回っています。クリック後の離脱を防ぐため、商品ページの改善や、まとめ買いクーポンの発行を推奨します。")
                                    else:
                                        st.info("【現状維持・要観察】極端な異常値は見られません。現在の運用ペースを維持しつつ、日々のROASの推移を監視してください。")
                            else:
                                st.info("ℹ️ 期間内の相関分析データが存在しません。")
                        else:
                            st.info("ℹ️ 期間内の楽天の売上データが存在しません。")
                    else:
                        st.info("ℹ️ 楽天の広告レポートがアップロードされていないため、相関分析を表示できません。")
                    
                    st.markdown("---")
                    
                    # --- 📊 広告費と売上の相関分析 (Yahoo!) ---
                    st.write("### 📊 広告費と売上の相関分析 (Yahoo!)")
                    if yahoo_ad_df is not None and not yahoo_ad_df.empty:
                        yahoo_ad_stores = sorted(yahoo_ad_df["店舗"].unique())
                        selected_y_store = st.selectbox(
                            "分析対象 of Yahoo!店舗を選択してください", 
                            yahoo_ad_stores,
                            key="yahoo_ad_store_selector"
                        )
                        
                        store_ad_df = yahoo_ad_df[yahoo_ad_df["店舗"] == selected_y_store]
                        
                        if selected_y_store == "Yahoo!":
                            yahoo_sales_daily = filtered_df[filtered_df["店舗名"].replace(["Yahoo1号店", "Yahoo2号店"], "OTHER").str.startswith("Yahoo", na=False)].groupby("日付")["売上合計値"].sum().reset_index()
                        else:
                            yahoo_sales_daily = filtered_df[filtered_df["店舗名"] == selected_y_store].groupby("日付")["売上合計値"].sum().reset_index()
                        
                        if not yahoo_sales_daily.empty:
                            store_ad_df_temp = store_ad_df.copy()
                            store_ad_df_temp["日付"] = pd.to_datetime(store_ad_df_temp["日付"]).dt.normalize()
                            yahoo_sales_daily["日付"] = pd.to_datetime(yahoo_sales_daily["日付"]).dt.normalize()
                            
                            y_scatter_data = pd.merge(yahoo_sales_daily, store_ad_df_temp[["日付", "広告費", "広告売上", "クリック数", "注文数"]], on="日付", how="left")
                            y_scatter_data["広告費"] = pd.to_numeric(y_scatter_data["広告費"], errors='coerce').fillna(0)
                            y_scatter_data["広告売上"] = pd.to_numeric(y_scatter_data["広告売上"], errors='coerce').fillna(0)
                            y_scatter_data["クリック数"] = pd.to_numeric(y_scatter_data["クリック数"], errors='coerce').fillna(0)
                            y_scatter_data["注文数"] = pd.to_numeric(y_scatter_data["注文数"], errors='coerce').fillna(0)
                            
                            if not y_scatter_data.empty:
                                st.markdown("""
                                    <style>
                                    [data-testid="stMetricLabel"] {
                                        width: 100% !important;
                                        min-width: 120px !important;
                                    }
                                    [data-testid="stMetricLabel"] > div,
                                    [data-testid="stMetricLabel"] p,
                                    [data-testid="stMetricLabel"] span,
                                    [data-testid="stMetricLabel"] label {
                                        white-space: normal !important;
                                        overflow: visible !important;
                                        text-overflow: clip !important;
                                        display: inline-block !important;
                                        min-width: 100px !important;
                                    }
                                    </style>
                                    """, unsafe_allow_html=True)
                                st.subheader(f"📊 広告パフォーマンスサマリー ({selected_y_store})")
                                y_total_cost = y_scatter_data["広告費"].sum()
                                y_total_sales = y_scatter_data["広告売上"].sum()
                                y_total_clicks = y_scatter_data["クリック数"].sum()
                                y_total_orders = y_scatter_data["注文数"].sum()
                                
                                y_roas = operator.mul(y_total_sales / y_total_cost, 100) if y_total_cost > 0 else 0.0
                                y_cpc = (y_total_cost / y_total_clicks) if y_total_clicks > 0 else 0.0
                                y_cvr = operator.mul(y_total_orders / y_total_clicks, 100) if y_total_clicks > 0 else 0.0
                                y_cpa = (y_total_cost / y_total_orders) if y_total_orders > 0 else 0.0
                                
                                y_cols_top = st.columns(3)
                                y_cols_top[0].metric("広告費", f"¥{int(y_total_cost):,}")
                                y_cols_top[1].metric("広告売上", f"¥{int(y_total_sales):,}")
                                y_cols_top[2].metric("ROAS", f"{y_roas:.1f}%")
                                
                                y_cols_bottom = st.columns(3)
                                y_cols_bottom[0].metric("CVR", f"{y_cvr:.1f}%")
                                y_cols_bottom[1].metric("CPC", f"¥{y_cpc:.1f}")
                                y_cols_bottom[2].metric("CPA", f"¥{int(y_cpa):,}")
                                
                                import numpy as np
                                y_x = y_scatter_data["広告費"].values
                                y_y = y_scatter_data["売上合計値"].values
                                
                                y_fig_scatter = go.Figure()
                                y_fig_scatter.add_trace(go.Scatter(
                                    x=y_x, y=y_y,
                                    mode='markers',
                                    name='日別実績',
                                    marker=dict(size=10, color='#ffc107', opacity=0.8, line=dict(width=1, color='white')),
                                    text=y_scatter_data["日付"].dt.strftime('%Y-%m-%d'),
                                    hovertemplate="<b>日付: %{text}</b><br>広告費: ¥%{x:,.0f}<br>総売上: ¥%{y:,.0f}<extra></extra>"
                                ))
                                
                                if len(y_x) > 1 and np.var(y_x) > 0:
                                    try:
                                        y_slope, y_intercept = np.polyfit(y_x, y_y, 1)
                                        y_x_line = np.array([np.min(y_x), np.max(y_x)])
                                        y_y_line = operator.mul(y_slope, y_x_line) + y_intercept
                                        y_fig_scatter.add_trace(go.Scatter(
                                            x=y_x_line, y=y_y_line,
                                            mode='lines',
                                            name='傾向線 (回帰直線)',
                                            line=dict(color='#ff00ff', width=2, dash='dash'),
                                            hovertemplate="傾向線<extra></extra>"
                                        ))
                                    except Exception:
                                        pass
                                
                                y_fig_scatter.update_layout(
                                    title=f"広告費と総売上の相関分布 ({selected_y_store})",
                                    template="plotly_dark",
                                    xaxis=dict(title="広告費 (円)", tickformat=",d"),
                                    yaxis=dict(title="総売上 (円)", tickformat=",d"),
                                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                                    height=450, margin=dict(l=20, r=20, t=60, b=20)
                                )
                                st.plotly_chart(y_fig_scatter, use_container_width=True)
                                
                                y_correlation = y_scatter_data["広告費"].corr(y_scatter_data["売上合計値"])
                                if pd.notna(y_correlation):
                                    st.caption(f"💡 広告費と総売上の相関係数 ({selected_y_store}): {y_correlation:.2f} （1.0に近いほど強い正の相関があります。目安: 0.7以上で強い相関、0.4〜0.7で中程度の相関）")
                                    st.caption("※紫色の点線は『傾向線（回帰直線）』です。この線の傾きが右肩上がりであるほど、広告費の増減が売上に直接影響を与えている（広告が効果的に機能している）ことを示します。逆に線が横ばいや右下がりの場合は、広告費と売上の連動性が低く、運用の見直しが必要なサインとなります。")
                                    
                                    st.subheader(f"💡 システムからの戦術提案 ({selected_y_store})")
                                    target_roas = yahoo_ad_summary.get(selected_y_store, {}).get("roas", 0.0)
                                    y_current_roas = target_roas if (target_roas is not None and pd.notna(target_roas)) else 0.0
                                    y_current_corr = y_correlation
                                    
                                    if y_current_corr >= 0.4 and y_current_roas >= 300.0:
                                        st.success("【投資推奨】広告と売上の連動性が高く、費用対効果も良好です。現在の着地予想を上回るために、日次予算の引き上げを検討してください。")
                                    elif y_current_corr < 0.2 and y_current_roas >= 300.0:
                                        st.warning("【予算削減・再配分推奨】広告のROASは高いですが、広告費と全体売上が連動していません。自然検索で売れるはずの注文を広告で獲得している（カニバリ）可能性があります。入札単価を下げて無駄なコストを削り、利益を確保してください。")
                                    elif y_current_corr < 0.2 and y_current_roas < 300.0:
                                        st.warning("【設定見直し推奨】広告費の増加が全体の売上増加に繋がっていません。無駄なクリックが発生している可能性があるため、アイテムマッチの除外商品設定や、費用対効果の高い他モールへの予算シフトを検討してください。")
                                    elif y_current_corr >= 0.4 and y_current_roas < 300.0:
                                        st.info("【CVR改善推奨】広告による集客効果は出ていますが、ROASが基準を下回っています。クリック後の離脱を防ぐため、商品ページの改善や、まとめ買いクーポンの発行を推奨します。")
                                    else:
                                        st.info("【現状維持・要観察】極端な異常値は見られません。現在の運用ペースを維持しつつ、日々のROASの推移を監視してください。")
                            else:
                                st.info(f"ℹ️ 期間内の {selected_y_store} の相関分析データが存在しません。")
                        else:
                            st.info(f"ℹ️ 期間内の {selected_y_store} の売上データが存在しません。")
                    else:
                        st.info("ℹ️ Yahoo!の広告レポートがアップロードされていないため、相関分析を表示できません。")
                    
                    st.markdown("---")
                    
                    # --- 🏆 商品別売上ランキング (Top 30) (Yahoo!セクション内) ---
                    st.write("#### 🏆 商品別売上ランキング (Top 30)")
                    st.caption("選択した店舗における「売上合計値（税込）」の上位30商品を表示します。")
                    
                    if yahoo_item_df is not None and not yahoo_item_df.empty:
                        available_item_stores = sorted(list(yahoo_item_df["店舗名"].unique()))
                        
                        col_store_sel, col_kw_sel = st.columns(2)
                        with col_store_sel:
                            selected_item_store = st.selectbox(
                                "分析対象の店舗を選択してください",
                                available_item_stores,
                                key="item_store_selector"
                            )
                        with col_kw_sel:
                            filter_keyword = st.text_input(
                                "絞り込みキーワード (任意)",
                                key="item_filter_keyword"
                            )
                        
                        df_store_all = yahoo_item_df[yahoo_item_df["店舗名"] == selected_item_store]
                        
                        df_store = df_store_all
                        target_month_display = "全期間"
                        if selected_month:
                            active_month_val = extract_month_from_str(selected_month)
                            df_store = df_store_all[df_store_all["対象月"] == active_month_val]
                            target_month_display = active_month_val
                        
                        if df_store.empty:
                            st.warning(f"⚠️ {selected_item_store} の {target_month_display} の商品データが存在しません。")
                        else:
                            if filter_keyword:
                                import unicodedata
                                norm_kw = unicodedata.normalize('NFKC', filter_keyword).lower()
                                name_normalized = df_store["商品名"].astype(str).str.normalize('NFKC').str.lower()
                                code_normalized = df_store["商品コード"].astype(str).str.normalize('NFKC').str.lower()
                                df_store = df_store[
                                    name_normalized.str.contains(norm_kw, na=False) |
                                    code_normalized.str.contains(norm_kw, na=False)
                                ]
                            
                            df_store_sorted = df_store.sort_values(by="売上合計値（税込）", ascending=False)
                            df_store_top30 = df_store_sorted.head(30)
                            
                            display_cols = ["店舗名", "商品コード", "商品名", "売上合計値（税込）", "注文点数"]
                            available_display_cols = [c for c in display_cols if c in df_store_top30.columns]
                            
                            df_store_top30_display = df_store_top30[available_display_cols].copy()
                            df_store_top30_display.index = range(1, len(df_store_top30_display) + 1)
                            
                            st.dataframe(df_store_top30_display, use_container_width=True)
                    else:
                        st.info("ℹ️ 商品データがロードされていないため、商品別売上ランキングを表示できません。")
                    
                    st.markdown("---")
                    
                    # --- 📊 広告費と売上の相関分析 (Amazon) ---
                    st.write("### 📊 広告費と売上の相関分析 (Amazon)")
                    if amazon_ad_df is not None and not amazon_ad_df.empty:
                        # 店舗選択セレクトボックスの追加
                        available_amz_stores = sorted(amazon_ad_df["店舗名"].unique())
                        selected_amz_store = st.selectbox("分析対象のAmazon店舗を選択してください", available_amz_stores, key="amz_store_select")
                        
                        amazon_ad_df_temp = amazon_ad_df.copy()
                        amazon_ad_df_temp["日付_normalized"] = pd.to_datetime(amazon_ad_df_temp["日付"], errors='coerce').dt.normalize()
                        amazon_filtered = amazon_ad_df_temp[(amazon_ad_df_temp["日付_normalized"] >= min_date) & (amazon_ad_df_temp["日付_normalized"] <= max_date) & (amazon_ad_df_temp["店舗名"] == selected_amz_store)]
                        
                        if not amazon_filtered.empty:
                            st.markdown("""
                                <style>
                                [data-testid="stMetricLabel"] {
                                    width: 100% !important;
                                    min-width: 120px !important;
                                }
                                [data-testid="stMetricLabel"] > div,
                                [data-testid="stMetricLabel"] p,
                                [data-testid="stMetricLabel"] span,
                                [data-testid="stMetricLabel"] label {
                                    white-space: normal !important;
                                    overflow: visible !important;
                                    text-overflow: clip !important;
                                    display: inline-block !important;
                                    min-width: 100px !important;
                                }
                                </style>
                                """, unsafe_allow_html=True)
                            
                            st.subheader(f"📊 広告パフォーマンスサマリー ({selected_amz_store})")
                            amz_total_cost = amazon_filtered["広告費"].sum()
                            amz_total_sales = amazon_filtered["広告売上"].sum()
                            amz_total_clicks = amazon_filtered["クリック数"].sum()
                            amz_total_orders = amazon_filtered["注文数"].sum()
                            
                            amz_roas = operator.mul(amz_total_sales / amz_total_cost, 100) if amz_total_cost > 0 else 0.0
                            amz_cpc = (amz_total_cost / amz_total_clicks) if amz_total_clicks > 0 else 0.0
                            amz_cvr = operator.mul(amz_total_orders / amz_total_clicks, 100) if amz_total_clicks > 0 else 0.0
                            amz_cpa = (amz_total_cost / amz_total_orders) if amz_total_orders > 0 else 0.0
                            
                            amz_cols_top = st.columns(3)
                            amz_cols_top[0].metric("広告費", f"¥{int(amz_total_cost):,}")
                            amz_cols_top[1].metric("広告売上", f"¥{int(amz_total_sales):,}")
                            amz_cols_top[2].metric("ROAS", f"{amz_roas:.1f}%")
                            
                            amz_cols_bottom = st.columns(3)
                            amz_cols_bottom[0].metric("CVR", f"{amz_cvr:.1f}%")
                            amz_cols_bottom[1].metric("CPC", f"¥{amz_cpc:.1f}")
                            amz_cols_bottom[2].metric("CPA", f"¥{int(amz_cpa):,}")
                            
                            import numpy as np
                            amz_x = amazon_filtered["広告費"].values
                            amz_y = amazon_filtered["商品売上"].values
                            
                            amz_fig_scatter = go.Figure()
                            amz_fig_scatter.add_trace(go.Scatter(
                                x=amz_x, y=amz_y,
                                mode='markers',
                                name='日別実績',
                                marker=dict(size=10, color='#ff9900', opacity=0.8, line=dict(width=1, color='white')),
                                text=amazon_filtered["日付"].dt.strftime('%Y-%m-%d'),
                                hovertemplate="<b>日付: %{text}</b><br>広告費: ¥%{x:,.0f}<br>商品売上: ¥%{y:,.0f}<extra></extra>"
                            ))
                            
                            if len(amz_x) > 1 and np.var(amz_x) > 0:
                                try:
                                    amz_slope, amz_intercept = np.polyfit(amz_x, amz_y, 1)
                                    amz_x_line = np.array([np.min(amz_x), np.max(amz_x)])
                                    amz_y_line = operator.mul(amz_slope, amz_x_line) + amz_intercept
                                    amz_fig_scatter.add_trace(go.Scatter(
                                        x=amz_x_line, y=amz_y_line,
                                        mode='lines',
                                        name='傾向線 (回帰直線)',
                                        line=dict(color='#ff00ff', width=2, dash='dash'),
                                        hovertemplate="傾向線<extra></extra>"
                                    ))
                                except Exception:
                                    pass
                            
                            amz_fig_scatter.update_layout(
                                title=f"広告費と商品売上の相関分布 ({selected_amz_store})",
                                template="plotly_dark",
                                xaxis=dict(title="広告費 (円)", tickformat=",d"),
                                yaxis=dict(title="商品売上 (円)", tickformat=",d"),
                                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                                height=450, margin=dict(l=20, r=20, t=60, b=20)
                            )
                            st.plotly_chart(amz_fig_scatter, use_container_width=True)
                            
                            amz_correlation = amazon_filtered["広告費"].corr(amazon_filtered["商品売上"])
                            if pd.notna(amz_correlation):
                                st.caption(f"💡 広告費と商品売上の相関係数 ({selected_amz_store}): {amz_correlation:.2f} （1.0に近いほど強い正の相関があります。目安: 0.7以上で強い相関、0.4〜0.7で中程度の相関）")
                                st.caption("※紫色の点線は『傾向線（回帰直線）』です。この線の傾きが右肩上がりであるほど、広告費の増減が売上に直接影響を与えている（広告が効果的に機能している）ことを示します。逆に線が横ばいや右下がりの場合は、広告費と売上の連動性が低く、運用の見直しが必要なサインとなります。")
                                
                                st.subheader(f"💡 システムからの戦術提案 ({selected_amz_store})")
                                amz_current_roas = amz_roas
                                amz_current_corr = amz_correlation
                                
                                if amz_current_corr >= 0.4 and amz_current_roas >= 300.0:
                                    st.success("【投資推奨】広告と売上の連動性が高く、費用対効果も良好です。現在の着地予想を上回るために、日次予算の引き上げを検討してください。")
                                elif amz_current_corr < 0.2 and amz_current_roas >= 300.0:
                                    st.warning("【予算削減・再配分推奨】広告のROASは高いですが、広告費と全体売上が連動していません。自然検索で売れるはずの注文を広告で獲得している（カニバリ）可能性があります。入札単価を下げて無駄なコストを削り、利益を確保してください。")
                                elif amz_current_corr < 0.2 and amz_current_roas < 300.0:
                                    st.warning("【設定見直し推奨】広告費の増加が全体の売上増加に繋がっていません。無駄なクリックが発生している可能性があるため、配信対象商品やキーワードのマッチタイプ見直し、費用対効果の高い他モールへの予算シフトを検討してください。")
                                elif amz_current_corr >= 0.4 and amz_current_roas < 300.0:
                                    st.info("【CVR改善推奨】広告による集客効果は出ていますが、ROASが基準を下回っています。商品詳細ページ（A+コンテンツなど）の改善や、プロモーション割引・クーポンの設定を推奨します。")
                                else:
                                    st.info("【現状維持・要観察】極端な異常値は見られません。現在の運用ペースを維持しつつ、日々のROAS of the storeを監視してください。")
                        else:
                            st.info(f"ℹ️ 期間内の {selected_amz_store} の相関分析データが存在しません。")
                    else:
                        st.info("ℹ️ Amazon of the store's advertising report is not uploaded.")
                
                with tab_items:
                    st.write("### 📦 商品横展開分析 ＆ 横展開チャンスの可視化")
                    
                    # アクティブ対象月の抽出
                    active_month_str = "不明"
                    if period_type == "月次" and selected_month:
                        active_month_str = extract_month_from_str(selected_month)
                    elif period_type == "週次" and selected_week:
                        active_month_str = extract_month_from_str(selected_week.split(" 〜 ")[0])
                    
                    if yahoo_item_df is None or yahoo_item_df.empty:
                        st.info("ℹ️ 商品データ（専用アップローダーからアップロードされたファイル）がアップロードされていないため、商品分析を表示できません。データ取り込みからアップロードしてください。")
                    else:
                        st.write("各店の商品別データから、売れ筋商品のランキングおよび店舗間での横展開推奨商品を自動的に抽出します。")
                        
                        # データが存在する全店舗の一覧を動的に取得
                        available_item_stores = sorted(list(yahoo_item_df["店舗名"].unique()))
                        
                        st.markdown("---")
                        st.write("#### 🏆 商品別売上ランキング (Top 30)")
                        st.caption("選択した店舗における「売上合計値（税込）」の上位30商品を表示します。")
                        
                        col_store_sel, col_kw_sel = st.columns(2)
                        with col_store_sel:
                            selected_item_store = st.selectbox(
                                "分析対象の店舗を選択してください",
                                available_item_stores,
                                key="item_store_selector_items_tab"
                            )
                        with col_kw_sel:
                            filter_keyword = st.text_input(
                                "絞り込みキーワード (任意)",
                                key="item_filter_keyword_items_tab"
                            )
                        
                        df_store_all = yahoo_item_df[yahoo_item_df["店舗名"] == selected_item_store]
                        df_store = df_store_all[df_store_all["対象月"] == active_month_str]
                        
                        if df_store.empty:
                            st.warning(f"⚠️ {selected_item_store} の {active_month_str} の商品データが存在しません。")
                        else:
                            if filter_keyword:
                                import unicodedata
                                norm_kw = unicodedata.normalize('NFKC', filter_keyword).lower()
                                name_normalized = df_store["商品名"].astype(str).str.normalize('NFKC').str.lower()
                                code_normalized = df_store["商品コード"].astype(str).str.normalize('NFKC').str.lower()
                                df_store = df_store[
                                    name_normalized.str.contains(norm_kw, na=False) |
                                    code_normalized.str.contains(norm_kw, na=False)
                                ]
                            
                            df_store_sorted = df_store.sort_values(by="売上合計値（税込）", ascending=False)
                            df_store_top30 = df_store_sorted.head(30)
                            
                            display_cols = ["店舗名", "商品コード", "商品名", "売上合計値（税込）", "注文点数"]
                            available_display_cols = [c for c in display_cols if c in df_store_top30.columns]
                            
                            df_store_top30_display = df_store_top30[available_display_cols].copy()
                            df_store_top30_display.index = range(1, len(df_store_top30_display) + 1)
                            
                            st.dataframe(df_store_top30_display, use_container_width=True)
                        
                        st.markdown("---")
                        st.write("#### 🚀 【横展開・強化推奨リスト】")
                        st.caption("選択した「成功店舗（基準）」で売上トップ50に入っているものの、「展開先店舗」では売上が立っていない、もしくはアクセスが極端に少ない商品を抽出し、横展開を推奨します。")
                        
                        col_sel1, col_sel2 = st.columns(2)
                        with col_sel1:
                            source_store = st.selectbox(
                                "成功店舗 (基準) を選択してください",
                                available_item_stores,
                                index=0,
                                key="source_store_selector"
                            )
                        with col_sel2:
                            target_store_options = [s for s in available_item_stores if s != source_store]
                            if not target_store_options:
                                target_store_options = available_item_stores
                            target_store = st.selectbox(
                                "展開先店舗 を選択してください",
                                target_store_options,
                                index=0,
                                key="target_store_selector"
                            )
                        
                        df_source_all = yahoo_item_df[yahoo_item_df["店舗名"] == source_store]
                        df_target_all = yahoo_item_df[yahoo_item_df["店舗名"] == target_store]
                        
                        df_source = df_source_all[df_source_all["対象月"] == active_month_str]
                        df_target = df_target_all[df_target_all["対象月"] == active_month_str]
                        
                        if df_source.empty:
                            st.warning(f"⚠️ {source_store} の {active_month_str} の商品データが不足しているため、横展開分析を実行できません。")
                        else:
                            df_source_top50 = df_source.sort_values(by="売上合計値（税込）", ascending=False).head(50)
                            
                            target_dict = {}
                            for _, row in df_target.iterrows():
                                code = str(row["商品コード"]).strip()
                                target_dict[code] = {
                                    "売上": row.get("売上合計値（税込）", 0.0),
                                    "商品名": row.get("商品名", "不明"),
                                    "注文点数": row.get("注文点数", 0.0)
                                }
                            
                            recommend_list = []
                            for _, row in df_source_top50.iterrows():
                                code = str(row["商品コード"]).strip()
                                source_sales = row.get("売上合計値（税込）", 0.0)
                                source_name = row.get("商品名", "不明")
                                source_orders = row.get("注文点数", 0.0)
                                
                                target_info = target_dict.get(code)
                                is_recommend = False
                                target_sales = 0.0
                                target_orders = 0.0
                                
                                if target_info is None:
                                    is_recommend = True
                                else:
                                    target_sales = target_info["売上"]
                                    target_orders = target_info["注文点数"]
                                    
                                    if target_sales == 0.0:
                                        is_recommend = True
                                
                                if is_recommend:
                                    recommend_list.append({
                                        "商品コード": code,
                                        "商品名": source_name,
                                        f"{source_store} 売上": source_sales,
                                        f"{source_store} 注文点数": source_orders,
                                        f"{target_store} 売上": target_sales,
                                        f"{target_store} 注文点数": target_orders
                                    })
                            
                            if recommend_list:
                                df_recommend = pd.DataFrame(recommend_list)
                                df_recommend.index = range(1, len(df_recommend) + 1)
                                st.success(f"💡 {source_store} の {active_month_str} の売れ筋から、 {target_store} へ未展開または改善の余地がある商品を {len(recommend_list)} 件検出しました。{target_store} への登録や広告強化をおすすめします。")
                                st.dataframe(df_recommend, use_container_width=True)
                                
                                # Excelファイルへの変換処理
                                buffer = io.BytesIO()
                                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                                    df_recommend.to_excel(writer, index=False)
                                excel_data = buffer.getvalue()
                                
                                st.download_button(
                                    label="📥 横展開リストをExcelでダウンロード",
                                    data=excel_data,
                                    file_name="cross_channel_expansion_list.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                            else:
                                st.info(f"🤝 現在、 {source_store} の {active_month_str} のトップ50商品はすべて {target_store} でも安定して売上または注文点数が立っています。")
                        
                        st.markdown("---")
                        st.write("#### 📈 季節商材の初動検知＆月別トレンド")
                        st.caption("月を跨いだ売上の急上昇（季節商材の初動）を自動検知します。")
                        
                        months_available = sorted([m for m in yahoo_item_df["対象月"].unique() if m != "不明"])
                        
                        if len(months_available) < 2:
                            st.info("ℹ️ トレンド分析を行うには、少なくとも2ヶ月分以上の商品データ（対象月が異なるデータ）が必要です。")
                        else:
                            col_m1, col_m2 = st.columns(2)
                            with col_m1:
                                base_month = st.selectbox("基準月を選択してください", months_available, index=0, key="base_month_selector")
                            with col_m2:
                                compare_month = st.selectbox("比較月を選択してください", months_available, index=min(1, len(months_available)-1), key="compare_month_selector")
                            
                            if base_month == compare_month:
                                st.warning("⚠️ 基準月と比較月には異なる月を選択してください。")
                            else:
                                df_base = yahoo_item_df[yahoo_item_df["対象月"] == base_month]
                                df_compare = yahoo_item_df[yahoo_item_df["対象月"] == compare_month]
                                
                                df_base_grouped = df_base.groupby(["店舗名", "商品コード"]).agg({
                                    "売上合計値（税込）": "sum",
                                    "商品名": "first"
                                }).reset_index()
                                
                                df_compare_grouped = df_compare.groupby(["店舗名", "商品コード"]).agg({
                                    "売上合計値（税込）": "sum",
                                    "注文点数": "sum",
                                    "商品名": "first"
                                }).reset_index()
                                
                                df_trend = pd.merge(
                                    df_base_grouped,
                                    df_compare_grouped,
                                    on=["店舗名", "商品コード"],
                                    how="outer",
                                    suffixes=("_基準", "_比較")
                                )
                                
                                df_trend["売上合計値（税込）_基準"] = df_trend["売上合計値（税込）_基準"].fillna(0.0)
                                df_trend["売上合計値（税込）_比較"] = df_trend["売上合計値（税込）_比較"].fillna(0.0)
                                df_trend["注文点数"] = df_trend["注文点数"].fillna(0.0)
                                df_trend["商品名"] = df_trend["商品名_比較"].fillna(df_trend["商品名_基準"]).fillna("不明")
                                
                                growth_rates = []
                                for idx, row in df_trend.iterrows():
                                    b_val = row["売上合計値（税込）_基準"]
                                    c_val = row["売上合計値（税込）_比較"]
                                    if b_val > 0:
                                        rate = operator.mul(c_val / b_val, 100.0)
                                    else:
                                        if c_val > 0:
                                            rate = 999.9
                                        else:
                                            rate = 0.0
                                    growth_rates.append(rate)
                                df_trend["売上成長率(%)"] = growth_rates
                                
                                df_target_items = df_trend[
                                    (df_trend["売上合計値（税込）_比較"] >= 10000.0) &
                                    (df_trend["売上成長率(%)"] >= 150.0)
                                ].copy()
                                
                                proposals = []
                                for idx, row in df_target_items.iterrows():
                                    rate = row["売上成長率(%)"]
                                    c_val = row["売上合計値（税込）_比較"]
                                    if rate >= 300.0:
                                        proposals.append("他モールでの広告予算引き上げ推奨")
                                    elif c_val >= 50000.0:
                                        proposals.append("在庫確保推奨")
                                    else:
                                        proposals.append("注力販売推奨（露出強化）")
                                df_target_items["戦術提案"] = proposals
                                
                                df_target_items = df_target_items.rename(columns={
                                    "売上合計値（税込）_基準": f"{base_month} 売上",
                                    "売上合計値（税込）_比較": f"{compare_month} 売上",
                                    "注文点数": f"{compare_month} 注文点数"
                                })
                                
                                display_cols_trend = [
                                    "店舗名",
                                    "商品コード",
                                    "商品名",
                                    f"{base_month} 売上",
                                    f"{compare_month} 売上",
                                    "売上成長率(%)",
                                    f"{compare_month} 注文点数",
                                    "戦術提案"
                                ]
                                
                                df_target_items_display = df_target_items[display_cols_trend].sort_values(by=f"{compare_month} 売上", ascending=False)
                                df_target_items_display.index = range(1, len(df_target_items_display) + 1)
                                
                                if df_target_items_display.empty:
                                    st.info(f"ℹ️ {base_month} から {compare_month} にかけて急上昇している商品は検出されませんでした。")
                                else:
                                    st.success(f"📈 {base_month} から {compare_month} にかけて売上が急上昇している商品を {len(df_target_items_display)} 件検出しました。")
                                    st.dataframe(
                                        df_target_items_display.style.format({
                                            f"{base_month} 売上": "¥{:,.0f}",
                                            f"{compare_month} 売上": "¥{:,.0f}",
                                            "売上成長率(%)": "{:.1f}%",
                                            f"{compare_month} 注文点数": "{:,.0f}"
                                        }),
                                        use_container_width=True
                                    )
                
            else: # 個別店舗
                available_stores = sorted(filtered_df["店舗名"].unique())
                selected_stores = st.sidebar.multiselect(
                    "店舗を選択", 
                    available_stores, 
                    default=[available_stores[0]] if available_stores else []
                )
                
                if selected_stores:
                    title_period = selected_month if period_type == "月次" else selected_week
                    st.write(f"### 🏪 選択店舗の統合分析 ({title_period})")
                    # 選択された店舗のデータを抽出
                    store_df_raw = filtered_df[filtered_df["店舗名"].isin(selected_stores)]
                    
                    # 日付でグループ化して合算
                    sum_cols = ["売上合計値", "前年売上", "月予算", "注文数合計"]
                    agg_dict = {c: "sum" for c in sum_cols if c in store_df_raw.columns}
                    
                    # イベント情報の結合（安全なロジック）
                    safe_join = lambda x: ", ".join([str(v) for v in x.unique() if pd.notna(v) and str(v).strip() not in ["", "nan", "None"]])
                    if "YSイベント" in store_df_raw.columns:
                        agg_dict["YSイベント"] = safe_join
                    if "ボーナスストア" in store_df_raw.columns:
                        agg_dict["ボーナスストア"] = safe_join
                    if "曜日" in store_df_raw.columns:
                        agg_dict["曜日"] = "first"
                    
                    store_df = store_df_raw.groupby("純粋な日付").agg(agg_dict).reset_index()
                    
                    # 合算後の再計算
                    store_df = store_df.sort_values("純粋な日付")
                    store_df["月予算"] = store_df["月予算"].max()
                    store_df["売上累計"] = store_df["売上合計値"].cumsum()
                    store_df["前年比"] = operator.mul(store_df["売上合計値"] / store_df["前年売上"], 100).replace([float('inf'), -float('inf')], 0).fillna(0)
                    store_df["予算比"] = operator.mul(store_df["売上累計"] / store_df["月予算"], 100).replace([float('inf'), -float('inf')], 0).fillna(0)
                    store_df["店舗名"] = ", ".join(selected_stores) if len(selected_stores) < 4 else f"{len(selected_stores)}店舗合算"
                    
                    cols = st.columns(2)
                    
                    this_store_sales = store_df['売上合計値'].sum() if not store_df.empty else 0
                    store_sales_delta = None
                    
                    if period_type == "週次" and last_week_start is not None:
                        # 選択された店舗の先週のデータを取得
                        last_week_store_df = combined_df[
                            (combined_df["店舗名"].isin(selected_stores)) & 
                            (combined_df["日付"] >= last_week_start) & 
                            (combined_df["日付"] <= last_week_end)
                        ]
                        last_week_store_sales = last_week_store_df["売上合計値"].sum() if not last_week_store_df.empty else 0
                        
                        if last_week_store_sales > 0:
                            store_sales_diff_pct = operator.mul((this_store_sales - last_week_store_sales) / last_week_store_sales, 100)
                            store_sales_delta = f"先週比 {store_sales_diff_pct:+.1f}% (先週: ¥{int(last_week_store_sales):,})"
                        else:
                            store_sales_delta = "先週比 N/A (先週実績なし)"
                    
                    cols[0].metric("選択店舗 合計売上", f"¥{int(this_store_sales):,}", delta=store_sales_delta)
                    cols[1].metric("予算進捗", f"{store_df['予算比'].iloc[-1] if not store_df.empty else 0:.1f}%")
                    
                    # --- 追加: 月末着地予想・目標逆算 ---
                    target_month = selected_month if period_type == "月次" else pd.to_datetime(selected_week.split(" 〜 ")[0]).strftime('%Y年%m月')
                    show_monthly_forecast_and_target_backcast(target_month, selected_stores_list=selected_stores)
                    
                    # --- 追加: 施策リフト分析 ---
                    show_event_lift_analysis(store_df)
                    
                    draw_plotly_chart(store_df, f"選択店舗分析: {', '.join(selected_stores[:3])}{'...' if len(selected_stores) > 3 else ''}")
                    st.dataframe(format_dataframe(store_df), use_container_width=True)
                else:
                    st.warning("⚠️ 分析する店舗を少なくとも1つ選択してください。")

else:
    st.info("👆 Yahoo!または楽天のExcelファイルをアップロードしてください。マルチモール分析が可能です。")

# フッター
st.markdown("---")
st.caption("© 2026 Cerberus Sync - Prototype v0.21 | Auto-Weekday Generation & Safety Fix")
