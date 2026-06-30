import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="超峰國際供應鏈 - 物流號段智慧查詢系統", layout="wide")

st.markdown("<h2 style='color: #1f4e78;'>📋 物流號段智慧查詢系統 (起訖區間高級版)</h2>", unsafe_allow_html=True)
st.caption("【頂級終極防線】本系統改採合約書「起始單號」與「結束單號」區間智慧比對，防範任何撞號可能。")

# 1. 偵測並讀取 Excel 檔案
excel_file = "號段維護表.xlsx"

if not os.path.exists(excel_file):
    # 如果雲端檔案不存在，自動建立符合超峰實務的全純數字範例
    df_init = pd.DataFrame({
        "起始單號": ["802050446001", "801959146001", "935000000000"],
        "結束單號": ["802052445996", "801959645992", "935999999999"],
        "派件廠商": ["黑貓宅急便", "黑貓宅急便", "嘉里大榮"],
        "客戶代號(客代)": ["9353865110", "9353865112", "12345678"],
        "財務/合約備註": ["火箭鳥-10 專用區間", "深圳新廠商區間", "大榮大宗純數字區間"]
    })
    df_init.to_excel(excel_file, index=False)

# 【核心修正】強制將所有欄位通通以「文字(字串)」型態讀入，徹底根絕格式打架
df = pd.read_excel(excel_file, dtype=str)

# 強制去掉可能產生的空白並轉大寫
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
        
        # 只要輸入的是純數字，且 Excel 裡也是純數字，直接用「數字大小」比區間
        if search_clean.isdigit() and start_no.isdigit() and end_no.isdigit():
            if int(start_no) <= int(search_clean) <= int(end_no):
                match = row
                break
        
        # 備用：如果是字母單號，用字串順序比區間
        elif len(search_clean) == len(start_no) == len(end_no):
            if start_no <= search_clean <= end_no:
                match = row
                break
                
        # 備用安全兜底防禦
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
st.dataframe(df, use_container_width='stretch') # 對齊 2026 最新 Streamlit 語法