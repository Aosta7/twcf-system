import streamlit as st
import pandas as pd
import os
import re
import requests
import hashlib
import json

# --- 頁面配置 ---
st.set_page_config(page_title="超峰國際供應鏈 - 物流號段與實時貨態系統", layout="wide")

st.markdown("<h2 style='color: #1f4e78;'>📋 物流號段與實時貨態智慧查詢系統</h2>", unsafe_allow_html=True)
st.caption("【企業旗艦版】已成功對接超峰後台 API，支援實時換單、下游派件商與完整物流軌跡查詢。")

# ================= 🔒 超峰 API 安全配置區 =================
API_URL = "http://cfcn.szcloudone.com/PlatForm/AllTransfer"
CF_CUS_ID = "20251089"      # 您提供的 CUSID
CF_MD5_KEY = "O8Rhdzwu"     # 您提供的加密 Key
CF_PLATFORM = 2014          # 您提供的 Client 代碼
# =========================================================

def get_chaofeng_logistics(bill_codes):
    """
    透過超峰 API 批次查詢單號的實時貨態與換單資訊
    """
    try:
        # 1. 組裝 JsonData 參數 (依文檔規範：PlatForm, CarrierBillCode, IsSign)
        json_data_dict = {
            "PlatForm": CF_PLATFORM,
            "CarrierBillCode": bill_codes,
            "IsSign": False
        }
        # 確保 JSON 字串緊湊無多餘空格，避免 MD5 算錯
        json_data_str = json.dumps(json_data_dict, separators=(',', ':'))
        
        # 2. 生成 MD5 簽名 (JsonData + Key 後轉大寫)
        sign_src = json_data_str + CF_MD5_KEY
        key_md5 = hashlib.md5(sign_src.encode('utf-8')).hexdigest().upper()
        
        # 3. 組裝外層公共請求參數
        payload = {
            "FunctionName": "GetOrderCarrierBillCodeLog",
            "JsonData": json_data_str,
            "CusID": CF_CUS_ID,
            "KeyMd5": key_md5
        }
        
        # 4. 發送 POST 請求
        response = requests.post(API_URL, json=payload, timeout=10)
        
        if response.status_code == 200:
            res_json = response.json()
            if res_json.get("State") == True:
                # 解析 ReturnJson 字串轉回 Python 字典
                return_data = json.loads(res_json.get("ReturnJson", "{}"))
                return return_data.get("list", [])
            else:
                st.error(f"❌ 超峰 API 回報錯誤: {res_json.get('Message')}")
    except Exception as e:
        st.error(f"💥 超峰 API 連線異常: {str(e)}")
    return []

# --- 1. 讀取 Excel 號段維護表（僅在後台運作，外人看不到） ---
excel_file = "號段維護表.xlsx"
if not os.path.exists(excel_file):
    df_init = pd.DataFrame({
        "起始單號": ["802050446001", "801959146001"],
        "結束單號": ["802052445996", "801959645992"],
        "派件廠商": ["黑貓宅急便", "黑貓宅急便"],
        "客戶代號(客代)": ["9353865110", "9353865112"],
        "黑貓API授權碼": ["Token_A", "Token_B"], 
        "財務/合約備註": ["火箭鳥區間", "深圳新廠商"]
    })
    df_init.to_excel(excel_file, index=False)

df = pd.read_excel(excel_file, dtype=str).fillna("")
df["起始單號"] = df["起始單號"].str.strip().str.upper()
df["結束單號"] = df["結束單號"].str.strip().str.upper()

# --- 2. 智慧查詢主介面 ---
st.write("---")
st.subheader("🔍 批次查單與實時貨態追蹤")

raw_input = st.text_area(
    "請貼入單號（每行一筆，或用逗號、空白隔開）：", 
    height=150, 
    placeholder="例如：\n802050446005\n801959146050"
)

if st.button("🚀 開始智慧識別並同步超峰貨態", type="primary"):
    if not raw_input.strip():
        st.warning("請先輸入至少一筆物流單號。")
    else:
        # 智慧解析輸入的單號清單
        input_list = [n.strip().upper() for n in re.split(r'[\n, \s]+', raw_input) if n.strip()]
        
        # --- 步驟 A：實時同步超峰 API 資料 ---
        with st.spinner("🔄 正在向超峰安全伺服器請求最新物流狀態..."):
            api_results = get_chaofeng_logistics(input_list)
        
        # 將 API 結果建立索引，方便用單號快速查找
        api_lookup = {item.get("CarrierBillCode"): item for item in api_results if item.get("CarrierBillCode")}
        
        # --- 步驟 B：號段比對與 API 資料交叉融合 ---
        results = []
        db_starts = df["起始單號"].values
        db_ends = df["結束單號"].values
        
        for search_clean in input_list:
            # 1. 進行本地 Excel 號段基礎過濾
            mask = ((df["起始單號"].str.len() == len(search_clean)) & (db_starts <= search_clean) & (search_clean <= db_ends))
            if not mask.any():
                mask = df["起始單號"].str.startswith(search_clean[:6])
            matched_rows = df[mask]
            
            # 預設合約內部的設定
            contract_vendor = matched_rows.iloc[0]["派件廠商"] if not matched_rows.empty else "未知區間"
            contract_remark = matched_rows.iloc[0]["財務/合約備註"] if not matched_rows.empty else "非合約常規單號"
            
            # 2. 檢查超峰 API 是否有回傳這筆單號的實時動態
            api_data = api_lookup.get(search_clean)
            
            if api_data:
                # 撈取超峰系統紀錄的最新狀態
                status_list = api_data.get("Status", [])
                
                if status_list:
                    # 預設拿最新的第一筆狀態
                    latest_status = status_list[0]
                    # 動態抓取超峰回傳的實際派件商（換單後），如果沒有就拿本地號段的派件商
                    realtime_vendor = latest_status.get("Logistics", api_data.get("CarrierName", contract_vendor))
                    
                    # 格式化顯示：時間 + 狀態 + 內容
                    current_status_desc = f"⏱️ [{latest_status.get('CreateTime')}] 【{latest_status.get('TypeName')}】{latest_status.get('Context')}"
                else:
                    realtime_vendor = api_data.get("CarrierName", contract_vendor)
                    current_status_desc = "📦 系統已建單，但目前下游物流商尚未更新軌跡"
                
                results.append({
                    "輸入單號": search_clean,
                    "系統識別": "✅ 成功連線",
                    "合約預期廠商": contract_vendor,
                    "實時派件廠商(換單後)": realtime_vendor,
                    "超峰系統實時貨態軌跡": current_status_desc,
                    "專案備註": contract_remark
                })
            else:
                # 本地號段有中，但超峰系統還沒建單的狀況
                results.append({
                    "輸入單號": search_clean,
                    "系統識別": "⚠️ 僅號段識別",
                    "合約預期廠商": contract_vendor,
                    "實時派件廠商(換單後)": contract_vendor if contract_vendor != "未知區間" else "-",
                    "超峰系統實時貨態軌跡": "❌ 超峰後台查無此單（請核對單號，或等待系統同步建檔）",
                    "專案備註": contract_remark
                })
        
        # --- 步驟 C：將大數據表格漂亮的呈現給使用者 ---
        res_df = pd.DataFrame(results)
        st.success(f"📋 查詢與實時同步完成！共處理 {len(input_list)} 筆單號。")
        st.dataframe(res_df, use_container_width=True)