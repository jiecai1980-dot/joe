# ============================================================
# Streamlit 資料整理工具
# 功能：上傳 Excel → 分離完整/不完整資料 → 下載結果
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np
import re
import io
from datetime import datetime
from openpyxl import Workbook

# ------------------------------------------------------------
# 頁面設定
# ------------------------------------------------------------
st.set_page_config(
    page_title="Excel 資料完整性檢查工具",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Excel 資料完整性檢查工具")
st.markdown("""
上傳 Excel 檔案後，系統會自動檢查資料完整性，並將結果分成兩個工作表：
- **完整資料**：通過所有檢查的乾淨資料
- **不完整資料**：含有缺失、格式錯誤、異常值等問題的資料
""")

# ------------------------------------------------------------
# 輔助函式：判斷不完整資料
# ------------------------------------------------------------

def is_missing(val):
    """判斷是否為缺失值或常見無效字串"""
    if pd.isna(val):
        return True
    s = str(val).strip().lower()
    invalid_list = [
        '', 'nan', 'none', 'null', 'n/a', 'na', '#na', '#n/a',
        'nil', 'unknown', '-', '--', 'n.a.', 'not available',
        'none_provided', 'invalid'
    ]
    return s in invalid_list


def is_valid_email(val):
    """簡單檢查 Email 格式"""
    if is_missing(val):
        return False
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(pattern, str(val).strip()))


def is_valid_phone(val):
    """檢查電話是否大致合理（至少含數字且長度合理）"""
    if is_missing(val):
        return False
    s = str(val).strip()
    digits = re.sub(r'[\s\-\+\(\)]', '', s)
    if not digits.isdigit():
        return False
    return 7 <= len(digits) <= 15


def is_valid_date(val):
    """嘗試解析日期，並排除明顯無效日期"""
    if is_missing(val):
        return False
    s = str(val).strip()

    if re.search(r'00-00|0000|02-31|02/31|2/31|1990-00-00', s):
        return False

    formats = [
        '%Y-%m-%d', '%Y/%m/%d', '%d/%m/%Y', '%m/%d/%Y',
        '%Y.%m.%d', '%d-%m-%Y', '%Y-%m-%d %H:%M:%S',
        '%Y/%m/%d %H:%M:%S', '%d/%m/%Y %H:%M:%S'
    ]
    for fmt in formats:
        try:
            datetime.strptime(s.split()[0], fmt)
            return True
        except Exception:
            continue
    return False


def is_abnormal_age(val):
    """年齡異常判斷"""
    if is_missing(val):
        return True
    try:
        age = float(str(val).strip())
        return age < 0 or age > 120
    except Exception:
        return True


def is_abnormal_salary(val):
    """薪資異常判斷（過低、過高或非數字）"""
    if is_missing(val):
        return True
    try:
        s = re.sub(r'[,$￥HKDUSD MOP]', '', str(val), flags=re.IGNORECASE)
        salary = float(s)
        return salary < 0 or salary > 500000
    except Exception:
        return True


def is_abnormal_credit(val):
    """信用評分異常（一般合理範圍約 300~850）"""
    if is_missing(val):
        return True
    try:
        score = float(str(val).strip())
        return score < 300 or score > 850
    except Exception:
        return True


def is_incomplete_row(row):
    """
    判斷一列是否為「不完整資料」
    只要符合任一重大問題，就歸類為不完整
    """
    # 1. 關鍵欄位缺失
    key_fields = [
        'First_Name', 'Last_Name', 'Email_Address', 'Phone_Number',
        'Department', 'Employment_Status', 'Salary_Monthly'
    ]
    for col in key_fields:
        if col in row.index and is_missing(row[col]):
            return True

    # 2. Email 格式錯誤
    if 'Email_Address' in row.index and not is_valid_email(row['Email_Address']):
        return True

    # 3. 電話格式錯誤
    if 'Phone_Number' in row.index and not is_valid_phone(row['Phone_Number']):
        return True

    # 4. 日期欄位檢查
    date_cols = [
        'DOB', 'Hire_Date', 'Contract_End_Date', 'Registration_Date',
        'Last_Login_Timestamp', 'Last_Purchase_Date'
    ]
    for col in date_cols:
        if col in row.index and not is_missing(row[col]) and not is_valid_date(row[col]):
            return True

    # 5. 年齡異常
    if 'Age' in row.index and is_abnormal_age(row['Age']):
        return True

    # 6. 薪資異常
    if 'Salary_Monthly' in row.index and is_abnormal_salary(row['Salary_Monthly']):
        return True

    # 7. 信用評分異常
    if 'Credit_Score' in row.index and is_abnormal_credit(row['Credit_Score']):
        return True

    # 8. 其他明顯異常字串
    if 'Years_of_Experience' in row.index:
        exp = str(row['Years_of_Experience']).strip().lower()
        if exp in ['99', '-3'] or 'yrs' in exp:
            return True

    return False


# ------------------------------------------------------------
# 主程式流程
# ------------------------------------------------------------

uploaded_file = st.file_uploader(
    "請上傳 Excel 檔案（.xlsx 或 .xls）",
    type=["xlsx", "xls"],
    help="支援 .xlsx 與 .xls 格式"
)

if uploaded_file is not None:
    try:
        # 讀取資料
        df = pd.read_excel(uploaded_file)
        st.success(f"✅ 成功讀取檔案：**{uploaded_file.name}**")
        st.write(f"原始資料共 **{len(df)}** 筆，**{len(df.columns)}** 個欄位")

        # 顯示原始資料預覽
        with st.expander("📋 查看原始資料預覽（前 10 筆）", expanded=False):
            st.dataframe(df.head(10), use_container_width=True)

        # 開始檢查
        if st.button("🚀 開始檢查資料完整性", type="primary"):
            with st.spinner("正在檢查資料完整性，請稍候..."):
                incomplete_mask = df.apply(is_incomplete_row, axis=1)

                df_complete = df[~incomplete_mask].copy()
                df_incomplete = df[incomplete_mask].copy()

            # 顯示統計結果
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("原始資料筆數", len(df))
            with col2:
                st.metric("完整資料筆數", len(df_complete), delta=None)
            with col3:
                st.metric("不完整資料筆數", len(df_incomplete), delta=None)

            st.markdown("---")

            # 顯示完整 / 不完整資料預覽
            tab1, tab2 = st.tabs(["✅ 完整資料", "⚠️ 不完整資料"])

            with tab1:
                st.write(f"共 **{len(df_complete)}** 筆完整資料")
                if len(df_complete) > 0:
                    st.dataframe(df_complete, use_container_width=True)
                else:
                    st.info("沒有通過檢查的完整資料")

            with tab2:
                st.write(f"共 **{len(df_incomplete)}** 筆不完整資料")
                if len(df_incomplete) > 0:
                    st.dataframe(df_incomplete, use_container_width=True)
                else:
                    st.success("恭喜！沒有發現不完整資料")

            # 產生 Excel 並提供下載
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_complete.to_excel(writer, sheet_name="完整資料", index=False)
                df_incomplete.to_excel(writer, sheet_name="不完整資料", index=False)

            output.seek(0)

            st.download_button(
                label="⬇️ 下載整理後的 Excel 檔案",
                data=output,
                file_name="整理後資料_完整與不完整.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

            st.success("✅ 處理完成！請點擊上方按鈕下載結果。")
            st.markdown("""
            **工作表說明：**
            - 「完整資料」：通過所有檢查的乾淨資料
            - 「不完整資料」：含有缺失、格式錯誤、異常值等問題的資料
            """)

    except Exception as e:
        st.error(f"❌ 讀取或處理檔案時發生錯誤：{str(e)}")
        st.info("請確認上傳的是有效的 Excel 檔案，且欄位名稱符合預期。")

else:
    st.info("👆 請先上傳 Excel 檔案開始使用")

# ------------------------------------------------------------
# 側邊欄說明
# ------------------------------------------------------------
with st.sidebar:
    st.header("📖 使用說明")
    st.markdown("""
    1. 點擊上方上傳按鈕選擇 Excel 檔案
    2. 點擊「開始檢查資料完整性」
    3. 查看完整與不完整資料
    4. 下載整理後的 Excel 檔案

    ---
    **檢查項目包含：**
    - 關鍵欄位是否缺失
    - Email 格式是否正確
    - 電話號碼是否合理
    - 日期格式是否有效
    - 年齡、薪資、信用評分是否異常
    """)

    st.markdown("---")
    st.caption("Powered by Streamlit + Pandas")