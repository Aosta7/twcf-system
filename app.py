import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="超峰國際供應鏈 - 物流號段智慧查詢系統", layout="wide")

st.markdown("<h2 style='color: #1f4e78;'>📋 物流號段智慧批次查詢系統</h2>", unsafe_allow_html=True)
st.caption("【企業旗艦版】支援多筆單號同時貼入查詢。未來可透過後台 Excel 擴充各契客之黑貓 API 授權碼。")

# 1. 偵測並讀取 Excel 檔案
excel_file = "號段維護表.xlsx"

if not os.path.exists(excel_file):
    # 自動建立相容未來 API 欄位的全新 Excel 結構
    df_init = pd.DataFrame({
        "起始單號": ["802050446001", "801959146001", "935000000000"],
        "結束單號": ["802052445996", "801959645992", "935999999999"],
        "派件廠商": ["黑貓宅急便", "黑貓宅急便", "嘉里大榮"],
        "客戶代號(客代)": ["9353865110", "9353865112", "12345678"],
        "黑貓API授權碼": ["Token_Example_A", "Token_Example_B", ""], # 留空代表非黑貓或暫無API
        "財務/合約備註": ["火箭鳥-10 專用區間", "深圳新廠商區間", "大榮大宗純數字區間"]
    })
    df_init.to_excel(excel_file, index=False)

# 強制以字串讀取 Excel
df = pd.read_excel(excel_file, dtype=str).fillna("")
df["起始單號"] = df["起始單號"].str.strip().str.upper()
df["結束單號"] = df["結束單號"].str.strip().str.upper()

# 2. 客服多筆查詢區
st.write("---")
st.subheader("🔍 批次查單（支援多筆同時輸入）")

# 改用 text_area，讓客服可以換行貼入多筆單號
raw_input = st.text_area("請貼入單號（每行一筆，或用逗號、空白隔開）：", height=150, placeholder="例如：\n802050446005\n801959146050")

if st.button("🚀 開始批次識別並查單"):
    if raw_input.strip() == "":
        st.warning("請先輸入至少一筆物流單號。")
    else:
        # 智慧切開客服輸入的字串（支援換行、逗號、空白）
        import re
        input_list = [n.strip().upper() for n in re.split(r'[\n, \s]+', raw_input) if n.strip()]
        
        # 準備存放查詢結果的清單
        results = []
        
        for search_clean in input_list:
            match_row = None
            
            # 區間比對核心邏輯
            for idx, row in df.iterrows():
                start_no = row["起始單號"]
                end_no = row["結束單號"]
                
                if search_clean.isdigit() and start_no.isdigit() and end_no.isdigit():
                    if int(start_no) <= int(search_clean) <= int(end_no):
                        match_row = row
                        break
                elif len(search_clean) == len(start_no) == len(end_no):
                    if start_no <= search_clean <= end_no:
                        match_row = row
                        break
                elif search_clean.startswith(start_no[:6]):
                    match_row = row
                    break
            
            # 將結果打包
            if match_row is not None:
                results.append({
                    "輸入單號": search_clean,
                    "識別狀態": "✅ 成功",
                    "派件廠商": match_row["派件廠商"],
                    "客戶代號 (客代)": match_row["客戶代號(客代)"],
                    "雲端 API 密鑰狀態": "已就緒 (有Token)" if match_row["黑貓API授權碼"] else "未配置",
                    "合約備註": match_row["財務/合約備註"]
                })
            else:
                results.append({
                    "輸入單號": search_clean,
                    "識別狀態": "❌ 失敗",
                    "派件廠商": "查無此合約區間",
                    "客戶代號 (客代)": "-",
                    "雲端 API 密鑰狀態": "-",
                    "合約備註": "請核對單號是否正確，或通知財務補件"
                })
        
        # 將結果轉成 DataFrame 並呈現在網頁上
        res_df = pd.DataFrame(results)
        st.success(f"📋 批次識別完成！共處理 {len(input_list)} 筆單號。")
        
        # 漂亮的表格輸出
        st.dataframe(res_df, use_container_width=True)

# 3. 網頁呈現目前資料庫
st.write("---")
st.subheader("📌 目前雲端 Excel 資料庫內容 (號段維護表.xlsx)")
st.dataframe(df, use_container_width='stretch')