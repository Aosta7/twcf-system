import streamlit as st
import pandas as pd
import os
import re

# --- 頁面配置 ---
st.set_page_config(page_title="超峰國際供應鏈 - 物流號段智慧查詢系統", layout="wide")

st.markdown("<h2 style='color: #1f4e78;'>📋 物流號段智慧批次查詢系統</h2>", unsafe_allow_html=True)
st.caption("【企業旗艦版】優化效能版。支援前導零單號精準比對、大數據批次快速識別。")

# --- 1. 偵測並讀取 Excel 檔案 ---
excel_file = "號段維護表.xlsx"

if not os.path.exists(excel_file):
    # 自動建立相容未來 API 欄位的全新 Excel 結構 (預設單號確保格式一致)
    df_init = pd.DataFrame({
        "起始單號": ["802050446001", "801959146001", "935000000000"],
        "結束單號": ["802052445996", "801959645992", "935999999999"],
        "派件廠商": ["黑貓宅急便", "黑貓宅急便", "嘉里大榮"],
        "客戶代號(客代)": ["9353865110", "9353865112", "12345678"],
        "黑貓API授權碼": ["Token_Example_A", "Token_Example_B", ""], 
        "財務/合約備註": ["火箭鳥-10 專用區間", "深圳新廠商區間", "大榮大宗純數字區間"]
    })
    df_init.to_excel(excel_file, index=False)

# 強制以字串讀取 Excel 並清理空白
df = pd.read_excel(excel_file, dtype=str).fillna("")
for col in ["起始單號", "結束單號"]:
    df[col] = df[col].str.strip().str.upper()

# --- 2. 客服多筆查詢區 ---
st.write("---")
st.subheader("🔍 批次查單（支援多筆同時輸入）")

raw_input = st.text_area(
    "請貼入單號（每行一筆，或用逗號、空白隔開）：", 
    height=150, 
    placeholder="例如：\n802050446005\n801959146050"
)

if st.button("🚀 開始批次識別並查單", type="primary"):
    if not raw_input.strip():
        st.warning("請先輸入至少一筆物流單號。")
    else:
        # 智慧切開客服輸入的字串
        input_list = [n.strip().upper() for n in re.split(r'[\n, \s]+', raw_input) if n.strip()]
        
        results = []
        
        # 預先提取維護表的向量陣列，加速比對速度
        db_starts = df["起始單號"].values
        db_ends = df["結束單號"].values
        
        # 核心優化：比對邏輯
        for search_clean in input_list:
            match_idx = None
            
            # 策略 1：完全長度與字串區間比對 (保留前導零防護)
            # 在 Python/Pandas 中，同長度的數字字串字元比對 (例如 "801" <= "802" <= "803") 與數字大小邏輯完全一致
            mask = (
                (df["起始單號"].str.len() == len(search_clean)) & 
                (db_starts <= search_clean) & 
                (search_clean <= db_ends)
            )
            
            # 策略 2：如果策略 1 沒對應到，啟動「前6碼模糊前綴比對」
            if not mask.any():
                search_prefix = search_clean[:6]
                # 檢查起始號段前6碼是否相同
                mask = df["起始單號"].str.startswith(search_prefix)
                
            # 取得第一個匹配成功的索引
            matched_rows = df[mask]
            
            if not matched_rows.empty:
                # 撈取第一筆匹配到的資料
                match_row = matched_rows.iloc[0]
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
        
        # 呈現結果
        res_df = pd.DataFrame(results)
        st.success(f"📋 批次識別完成！共處理 {len(input_list)} 筆單號。")
        
        # 透過 st.dataframe 的 column_config 讓 UI 看起來更專業
        st.dataframe(
            res_df, 
            use_container_width=True,
            column_config={
                "識別狀態": st.column_config.TextColumn("識別狀態", width="small"),
                "輸入單號": st.column_config.TextColumn("輸入單號", help="客服輸入的原始單號"),
            }
        )

# --- 3. 網頁呈現目前資料庫 ---
st.write("---")
st.subheader("📌 目前雲端 Excel 資料庫內容 (號段維護表.xlsx)")

# 提供下載按鈕讓客服或財務可以直接下載這個 Excel 範本進行修改
with open(excel_file, "rb") as f:
    st.download_button(
        label="📥 下載目前號段維護表 (Excel)",
        data=f,
        file_name="號段維護表.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.dataframe(df, use_container_width=True)