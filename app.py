# ============================================
# Streamlit - 資料品質整理與自動下載
# （含完整 / 不完整資料統計摘要）
# ============================================

import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime
from collections import Counter
import io
import warnings

warnings.filterwarnings("ignore")

# --------------------------------------------
# Page config
# --------------------------------------------
st.set_page_config(
    page_title="資料品質整理工具",
    page_icon="📊",
    layout="wide",
)

st.title("📊 資料品質整理與分類工具")
st.markdown(
    "上傳 Excel 檔案（.xlsx），自動檢查資料品質，並輸出完整 / 不完整資料與統計摘要。"
)

# --------------------------------------------
# 資料品質檢查函式（與原 Colab 相同）
# --------------------------------------------

def is_null_like(value):
    """判斷是否為空值 / NULL / N/A / #NA 等"""
    if pd.isna(value):
        return True
    if isinstance(value, str):
        v = value.strip().upper()
        null_keywords = [
            "",
            "NULL",
            "N/A",
            "#N/A",
            "#NA",
            "NA",
            "NONE",
            "NONE_PROVIDED",
            "NIL",
            "UNKNOWN",
            "-",
            "--",
            "NAN",
            "NAT",
            "INVALID",
            "INVALID_TIME",
            "INVALID_MAC",
        ]
        return v in null_keywords or v.startswith("INVALID")
    return False


def is_valid_date(value):
    """檢查日期格式是否合理"""
    if is_null_like(value):
        return False
    value_str = str(value).strip()

    if re.search(r"00-00|0000|INVALID|1990-00-00|2023-02-31", value_str, re.IGNORECASE):
        return False

    formats = [
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%Y.%m.%d",
        "%d-%m-%Y",
        "%Y-%m-%d %H:%M:%S",
    ]
    for fmt in formats:
        try:
            dt = datetime.strptime(value_str[:10], fmt)
            if 1950 <= dt.year <= 2030:
                return True
        except Exception:
            continue
    return False


def is_valid_phone(value):
    """檢查電話號碼格式"""
    if is_null_like(value):
        return False
    value_str = str(value).strip()

    if len(value_str) < 8:
        return False
    if value_str in ["12345", "6612-3456"]:
        return False

    cleaned = re.sub(r"[\s\-\(\)\+]", "", value_str)
    if not cleaned.isdigit():
        return False
    return True


def is_valid_email(value):
    """檢查 Email 格式"""
    if is_null_like(value):
        return False
    value_str = str(value).strip()
    pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return bool(re.match(pattern, value_str))


def is_valid_age(value):
    """檢查年齡是否合理"""
    if is_null_like(value):
        return False
    try:
        age = float(str(value).strip())
        return 16 <= age <= 80
    except Exception:
        return False


def is_valid_salary(value):
    """檢查薪資是否合理"""
    if is_null_like(value):
        return False
    try:
        cleaned = re.sub(r"[^\d\.\-]", "", str(value))
        salary = float(cleaned)
        return 5000 <= salary <= 200000
    except Exception:
        return False


def is_valid_credit_score(value):
    """檢查信用分數（一般 300~850）"""
    if is_null_like(value):
        return False
    try:
        score = float(str(value).strip())
        return 300 <= score <= 850
    except Exception:
        return False


def check_row_quality(row):
    """
    檢查單列資料品質
    回傳 (is_complete: bool, issues: list)
    """
    issues = []

    # 1. 關鍵欄位空值檢查
    key_fields = [
        "First_Name",
        "Last_Name",
        "Email_Address",
        "Phone_Number",
        "Department",
        "Employment_Status",
        "Salary_Monthly",
    ]
    for field in key_fields:
        if field in row.index and is_null_like(row[field]):
            issues.append(f"{field} 為空值")

    # 2. 年齡
    if "Age" in row.index and not is_valid_age(row["Age"]):
        issues.append("Age 異常或格式錯誤")

    # 3. 日期欄位
    date_fields = [
        "DOB",
        "Hire_Date",
        "Contract_End_Date",
        "Registration_Date",
        "Last_Login_Timestamp",
        "Last_Purchase_Date",
    ]
    for field in date_fields:
        if field in row.index and not is_null_like(row[field]):
            if not is_valid_date(row[field]):
                issues.append(f"{field} 日期格式錯誤")

    # 4. 電話
    if "Phone_Number" in row.index and not is_valid_phone(row["Phone_Number"]):
        issues.append("Phone_Number 格式錯誤")

    # 5. Email
    if "Email_Address" in row.index and not is_valid_email(row["Email_Address"]):
        issues.append("Email_Address 格式錯誤")

    # 6. 薪資
    if "Salary_Monthly" in row.index and not is_valid_salary(row["Salary_Monthly"]):
        issues.append("Salary_Monthly 異常值")

    # 7. 信用分數
    if "Credit_Score" in row.index and not is_valid_credit_score(row["Credit_Score"]):
        issues.append("Credit_Score 異常值")

    # 8. 消費金額異常
    if "Total_Spend_USD" in row.index:
        try:
            spend = float(re.sub(r"[^\d\.\-]", "", str(row["Total_Spend_USD"])))
            if spend < 0 or spend > 100000:
                issues.append("Total_Spend_USD 異常值")
        except Exception:
            if not is_null_like(row["Total_Spend_USD"]):
                issues.append("Total_Spend_USD 格式錯誤")

    # 9. 就業狀態大小寫 / 異常值
    if "Employment_Status" in row.index:
        status = (
            str(row["Employment_Status"]).strip()
            if not is_null_like(row["Employment_Status"])
            else ""
        )
        if status and status not in [
            "Active",
            "Pending",
            "Terminated",
            "Inactive",
            "Probation",
        ]:
            if status.lower() not in [
                "active",
                "pending",
                "terminated",
                "inactive",
                "probation",
            ]:
                issues.append("Employment_Status 值異常")
            elif (
                status != status.capitalize()
                and status != status.upper()
                and status != status.lower()
            ):
                issues.append("Employment_Status 大小寫混亂")

    is_complete = len(issues) == 0
    return is_complete, issues


def build_summary(df, df_complete, df_incomplete, df_issues):
    """建立統計摘要 DataFrame"""
    total_rows = len(df)
    complete_count = len(df_complete)
    incomplete_count = len(df_incomplete)

    summary_data = [
        ["項目", "數值", "說明"],
        ["原始資料總筆數", total_rows, "上傳檔案的總列數"],
        ["完整資料筆數", complete_count, "通過所有品質檢查的資料"],
        ["不完整資料筆數", incomplete_count, "有任何品質問題的資料"],
        [
            "完整資料佔比 (%)",
            round(complete_count / total_rows * 100, 2) if total_rows > 0 else 0,
            "完整資料 / 總筆數",
        ],
        [
            "不完整資料佔比 (%)",
            round(incomplete_count / total_rows * 100, 2) if total_rows > 0 else 0,
            "不完整資料 / 總筆數",
        ],
        ["", "", ""],
    ]

    # 問題類型統計
    if not df_issues.empty:
        all_issues = []
        for issues_str in df_issues["Issues"]:
            all_issues.extend([i.strip() for i in issues_str.split("|")])

        issue_counter = Counter(all_issues)

        summary_data.append(["=== 問題類型統計 ===", "", ""])
        summary_data.append(["問題類型", "出現次數", "佔不完整資料比例 (%)"])

        for issue, count in issue_counter.most_common():
            pct = (
                round(count / incomplete_count * 100, 2) if incomplete_count > 0 else 0
            )
            summary_data.append([issue, count, pct])
    else:
        summary_data.append(["=== 問題類型統計 ===", "", ""])
        summary_data.append(["無任何問題", 0, 0])

    summary_data.append(["", "", ""])
    summary_data.append(["=== 欄位空值統計（原始資料） ===", "", ""])
    summary_data.append(["欄位名稱", "空值筆數", "空值比例 (%)"])

    for col in df.columns:
        null_like_cnt = df[col].apply(is_null_like).sum()
        if null_like_cnt > 0:
            pct = round(null_like_cnt / total_rows * 100, 2)
            summary_data.append([col, int(null_like_cnt), pct])

    return pd.DataFrame(summary_data[1:], columns=summary_data[0])


def generate_excel(df_complete, df_incomplete, df_issues, df_summary):
    """產生多工作表 Excel bytes"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_complete.to_excel(writer, sheet_name="完整資料_Complete", index=False)
        df_incomplete.to_excel(writer, sheet_name="不完整資料_Incomplete", index=False)
        if not df_issues.empty:
            df_issues.to_excel(writer, sheet_name="問題摘要_Issues", index=False)
        df_summary.to_excel(writer, sheet_name="統計摘要_Summary", index=False)
    output.seek(0)
    return output.getvalue()


# --------------------------------------------
# Sidebar - 說明
# --------------------------------------------
with st.sidebar:
    st.header("使用說明")
    st.markdown(
        """
        1. 點擊下方「Browse files」上傳 `.xlsx` 檔案  
        2. 系統會自動執行資料品質檢查  
        3. 查看統計摘要與問題明細  
        4. 下載整理後的 Excel（含 4 個工作表）

        **檢查項目包含：**
        - 關鍵欄位空值
        - 年齡合理性
        - 日期格式
        - 電話 / Email 格式
        - 薪資 / 信用分數範圍
        - 就業狀態標準值
        """
    )
    st.markdown("---")
    st.caption("原 Colab 版本轉換為 Streamlit")

# --------------------------------------------
# 主畫面：上傳與處理
# --------------------------------------------
uploaded_file = st.file_uploader(
    "請上傳 Excel 檔案（.xlsx）",
    type=["xlsx"],
    help="支援 .xlsx 格式",
)

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.success(f"已成功讀取檔案：**{uploaded_file.name}**")
        st.info(f"原始資料共 **{len(df)}** 筆，**{len(df.columns)}** 個欄位")

        with st.spinner("正在進行資料品質檢查..."):
            complete_rows = []
            incomplete_rows = []
            issue_details = []

            for idx, row in df.iterrows():
                is_complete, issues = check_row_quality(row)

                if is_complete:
                    complete_rows.append(row)
                else:
                    incomplete_rows.append(row)
                    issue_details.append(
                        {
                            "Record_ID": row.get("Record_ID", idx),
                            "Issues": " | ".join(issues),
                        }
                    )

            df_complete = pd.DataFrame(complete_rows)
            df_incomplete = pd.DataFrame(incomplete_rows)
            df_issues = pd.DataFrame(issue_details)
            df_summary = build_summary(df, df_complete, df_incomplete, df_issues)

        # ---------- 指標卡片 ----------
        col1, col2, col3, col4 = st.columns(4)
        total = len(df)
        complete_cnt = len(df_complete)
        incomplete_cnt = len(df_incomplete)
        complete_pct = round(complete_cnt / total * 100, 2) if total > 0 else 0

        col1.metric("原始總筆數", total)
        col2.metric("完整資料", complete_cnt, f"{complete_pct}%")
        col3.metric("不完整資料", incomplete_cnt, f"{100 - complete_pct:.2f}%")
        col4.metric("欄位數", len(df.columns))

        st.markdown("---")

        # ---------- Tabs ----------
        tab1, tab2, tab3, tab4 = st.tabs(
            [
                "✅ 完整資料",
                "⚠️ 不完整資料",
                "📋 問題摘要",
                "📈 統計摘要",
            ]
        )

        with tab1:
            st.subheader(f"完整資料（{len(df_complete)} 筆）")
            if not df_complete.empty:
                st.dataframe(df_complete, use_container_width=True, height=400)
            else:
                st.warning("沒有通過品質檢查的完整資料")

        with tab2:
            st.subheader(f"不完整資料（{len(df_incomplete)} 筆）")
            if not df_incomplete.empty:
                st.dataframe(df_incomplete, use_container_width=True, height=400)
            else:
                st.success("所有資料皆完整，沒有不完整資料！")

        with tab3:
            st.subheader("問題摘要（每筆不完整資料的問題原因）")
            if not df_issues.empty:
                st.dataframe(df_issues, use_container_width=True, height=400)
            else:
                st.success("無任何問題")

        with tab4:
            st.subheader("統計摘要")
            st.dataframe(df_summary, use_container_width=True, height=500)

        # ---------- 下載 ----------
        st.markdown("---")
        st.subheader("📥 下載整理後的 Excel")

        excel_bytes = generate_excel(df_complete, df_incomplete, df_issues, df_summary)
        output_filename = "整理後_資料品質分類.xlsx"

        st.download_button(
            label="⬇️ 下載 整理後_資料品質分類.xlsx",
            data=excel_bytes,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        st.caption(
            "工作表內容：1. 完整資料_Complete　2. 不完整資料_Incomplete　"
            "3. 問題摘要_Issues　4. 統計摘要_Summary"
        )

    except Exception as e:
        st.error(f"讀取或處理檔案時發生錯誤：{e}")
        st.exception(e)
else:
    st.info("👆 請先上傳 Excel 檔案以開始資料品質檢查")