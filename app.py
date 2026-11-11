import os
import re

import pandas as pd
import streamlit as st

# --------------------------------------------------
# 頁面設定
# --------------------------------------------------
st.set_page_config(
    page_title="超峰國際供應鏈 - 物流號段查詢系統",
    page_icon="📦",
    layout="wide",
)

st.markdown(
    "<h2 style='color:#1f4e78;margin-bottom:0.25rem;'>📦 物流號段智慧批次查詢系統</h2>",
    unsafe_allow_html=True,
)
st.caption("公司內部使用｜支援黑貓宅急便、嘉里大榮｜可一次貼入多筆單號")

# --------------------------------------------------
# 可選：Streamlit Secrets 設定 APP_PASSWORD 後啟用密碼保護
# 公開 GitHub repo 請勿把密碼或 API Token 寫進程式碼 / Excel
# --------------------------------------------------
def get_optional_secret(name: str) -> str:
    try:
        return str(st.secrets.get(name, "")).strip()
    except Exception:
        return ""


app_password = get_optional_secret("APP_PASSWORD")
if app_password:
    password_input = st.text_input("內部查詢密碼", type="password")
    if password_input != app_password:
        if password_input:
            st.error("密碼錯誤。")
        st.stop()

# --------------------------------------------------
# 資料讀取
# --------------------------------------------------
EXCEL_FILE = "號段維護表.xlsx"
REQUIRED_COLUMNS = [
    "起始單號",
    "結束單號",
    "派件廠商",
    "內部代碼",
    "所屬客戶",
    "來源客代",
    "實際使用客代",
    "數量",
    "狀態",
    "用途/備註",
]


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    # 避免 Excel 數字欄被讀成 12345.0
    if re.fullmatch(r"\d+\.0", text):
        return text[:-2]
    return text


@st.cache_data(show_spinner=False)
def load_ranges(file_path: str, modified_time: float) -> pd.DataFrame:
    # modified_time 只用來讓檔案更新後 cache 自動失效
    _ = modified_time
    data = pd.read_excel(file_path, dtype=str).fillna("")
    data.columns = [str(c).strip() for c in data.columns]

    missing = [c for c in REQUIRED_COLUMNS if c not in data.columns]
    if missing:
        raise ValueError("號段維護表缺少欄位：" + "、".join(missing))

    for col in REQUIRED_COLUMNS:
        data[col] = data[col].map(clean_text)

    data = data[(data["起始單號"] != "") & (data["結束單號"] != "")].copy()
    return data


if not os.path.exists(EXCEL_FILE):
    st.error(f"找不到 `{EXCEL_FILE}`，請確認 Excel 已與 app.py 放在同一個 GitHub repo 根目錄。")
    st.stop()

try:
    df = load_ranges(EXCEL_FILE, os.path.getmtime(EXCEL_FILE))
except Exception as exc:
    st.error(f"讀取號段維護表失敗：{exc}")
    st.stop()

# --------------------------------------------------
# 驗證 / 查詢函式
# --------------------------------------------------
def blackcat_check_digit_is_valid(number: str) -> bool:
    """黑貓 12 碼：第 12 碼 = 前 11 碼各位數字總和 mod 7。"""
    if not re.fullmatch(r"\d{12}", number):
        return False
    expected = sum(int(d) for d in number[:11]) % 7
    return int(number[-1]) == expected


def find_range(number: str):
    if not number.isdigit():
        return None

    value = int(number)
    for _, row in df.iterrows():
        start = clean_text(row["起始單號"])
        end = clean_text(row["結束單號"])

        if not (start.isdigit() and end.isdigit()):
            continue
        if len(number) != len(start):
            continue
        if int(start) <= value <= int(end):
            return row
    return None


def display_code(row) -> str:
    actual = clean_text(row["實際使用客代"])
    source = clean_text(row["來源客代"])
    return actual or source or "-"


# --------------------------------------------------
# 查詢介面
# --------------------------------------------------
st.write("---")
st.subheader("🔎 批次查單")

raw_input = st.text_area(
    "請貼入單號（每行一筆，也支援逗號、空白分隔）：",
    height=170,
    placeholder="例如：\n802104746004\n40558512701",
)

col_btn, col_info = st.columns([1, 5])
with col_btn:
    search_clicked = st.button("開始批次查詢", type="primary", use_container_width=True)
with col_info:
    st.caption(f"目前維護 {len(df):,} 個號段區間")

if search_clicked:
    if not raw_input.strip():
        st.warning("請先輸入至少一筆物流單號。")
        st.stop()

    input_list = [
        n.strip()
        for n in re.split(r"[\s,，;；]+", raw_input)
        if n.strip()
    ]

    results = []

    for number in input_list:
        row = find_range(number)

        if row is None:
            results.append(
                {
                    "輸入單號": number,
                    "識別狀態": "❌ 查無號段",
                    "派件廠商": "-",
                    "所屬客戶": "-",
                    "內部代碼": "-",
                    "來源客代": "-",
                    "實際使用客代": "-",
                    "號段狀態": "-",
                    "號段範圍": "-",
                    "備註": "請核對單號，或通知管理員更新號段維護表",
                }
            )
            continue

        carrier = clean_text(row["派件廠商"])
        identification = "✅ 成功"
        note = clean_text(row["用途/備註"])

        # 落在黑貓區間，但 12 碼檢查碼錯誤時額外提醒
        if carrier == "黑貓宅急便" and len(number) == 12 and not blackcat_check_digit_is_valid(number):
            identification = "⚠️ 區間吻合／檢查碼異常"
            check_note = "輸入值落在此號段，但黑貓第 12 碼檢查碼不符，請再核對單號"
            note = f"{check_note}｜{note}" if note else check_note

        results.append(
            {
                "輸入單號": number,
                "識別狀態": identification,
                "派件廠商": carrier,
                "所屬客戶": clean_text(row["所屬客戶"]) or "-",
                "內部代碼": clean_text(row["內部代碼"]) or "-",
                "來源客代": clean_text(row["來源客代"]) or "-",
                "實際使用客代": clean_text(row["實際使用客代"]) or "-",
                "號段狀態": clean_text(row["狀態"]) or "-",
                "號段範圍": f'{clean_text(row["起始單號"])} ～ {clean_text(row["結束單號"])}',
                "備註": note or "-",
            }
        )

    res_df = pd.DataFrame(results)

    success_count = int(res_df["識別狀態"].str.startswith("✅").sum())
    warning_count = int(res_df["識別狀態"].str.startswith("⚠️").sum())
    failed_count = int(res_df["識別狀態"].str.startswith("❌").sum())

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("查詢筆數", f"{len(res_df):,}")
    m2.metric("成功", f"{success_count:,}")
    m3.metric("需核對", f"{warning_count:,}")
    m4.metric("查無號段", f"{failed_count:,}")

    st.dataframe(
        res_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "輸入單號": st.column_config.TextColumn(width="medium"),
            "識別狀態": st.column_config.TextColumn(width="medium"),
            "派件廠商": st.column_config.TextColumn(width="small"),
            "所屬客戶": st.column_config.TextColumn(width="medium"),
            "內部代碼": st.column_config.TextColumn(width="small"),
            "來源客代": st.column_config.TextColumn(width="medium"),
            "實際使用客代": st.column_config.TextColumn(width="medium"),
            "號段狀態": st.column_config.TextColumn(width="small"),
            "號段範圍": st.column_config.TextColumn(width="large"),
            "備註": st.column_config.TextColumn(width="large"),
        },
    )

    csv_data = res_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "下載查詢結果 CSV",
        data=csv_data,
        file_name="物流號段查詢結果.csv",
        mime="text/csv",
    )
