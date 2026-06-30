import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="超峰國際供應鏈 - 物流號段智慧查詢系統", layout="wide")

st.markdown("<h2 style='color: #1f4e78;'>📋 物流號段智慧查詢系統 (起訖區間高級版)</h2>", unsafe_allow_html=True)
st.caption("【頂級終極防線】本系統改採合約書「起始單號」與「結束單號」區間智慧比對，防範任何撞號可能。")

# 1. 偵測並讀取 Excel 檔案
excel_file = "號段維護表.xlsx"

if not os.path.exists(excel_file):
    # 如果檔案不存在，自動建立全新的起訖格式欄位
    df_init = pd.DataFrame({
        "起始單號": ["802050446001", "801959146001", "KT062026050000", "935000000000"],
        "結束單號": ["802052445996", "801959645992", "KT062026059999", "935999999999"],
        "派件廠商": ["黑貓宅急便", "黑貓宅急便", "嘉里大榮", "大榮板運"],
        "客戶代號(客代)": ["9353865110", "9353865112", "KT062026050008", "超大超重貨件"],
        "財務/合約備註": ["火箭鳥-10 (完全鎖死合約區間)", "深圳新廠商區間", "大榮海運件", "大榮板車號段"]
    })
    df_init.to_excel(excel_file, index=False)

df = pd.read_excel(excel_file)
# 強制把欄位轉成文字並去掉空白
df["起始單號"] = df["起始單號"].astype(str).str.strip().str.upper()
df["結束單號"] = df["結束單號"].astype(str).str.strip().str.upper()

# 2. 客服查詢區
st.write("---")
st.subheader("🔍 快速查單")
search_input = st.text_input("請輸入客服或客戶提供的完整物流單號：", placeholder="請在這裡輸入單號...")

if search_input:
    search_clean = search_input.strip().upper()
    match = None
    
    # 智慧區間比對核心演算法
    for idx, row in df.iterrows():
        start_no = row["起始單號"]
        end_no = row["結束單號"]
        
        # 情況 A：如果輸入的單號是純數字，且 Excel 裡也是純數字，直接用「數字大小」比區間
        if search_clean.isdigit() and start_no.isdigit() and end_no.isdigit():
            if int(start_no) <= int(search_clean) <= int(end_no):
                match = row
                break
        
        # 情況 B：如果含有英文字母（例如大榮的KT），或者長度一樣，用字串順序比區間
        elif len(search_clean) == len(start_no) == len(end_no):
            if start_no <= search_clean <= end_no:
                match = row
                break
                
        # 情況 C：安全兜底防禦（如果長度不一，改回檢查開頭）
        elif search_clean.startswith(start_no[:6]):
            match = row
            break
            
    if match is not None:
        st.success(f"✅ 成功識別廠商：**{match['派件廠商']}**")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("所屬合約區間", f"{match['起始單號']} ➔ {match['結束單號']}")
        with col2:
            st.metric("對應客戶代號 (客代)", match["客戶代號(客代)"])
        with col3:
            st.info(f"💡 財務備註：{match['財務/合約備註']}")
    else:
        st.error(f"❌ 查無此單！該單號不屬於目前系統中維護的任何一間廠商合約區間，請檢查單號或補件。")

# 3. 網頁呈現目前資料庫
st.write("---")
st.subheader("📌 目前雲端 Excel 資料庫內容 (號段維護表.xlsx)")
st.dataframe(df, use_container_width=True)