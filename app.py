import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

# ==========================================================
# 🎨 THEME SYSTEM (Premium Minimal)
# ==========================================================

def apply_theme():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@300;400;600;700&display=swap');

    html, body, [class*="css"]  {
        font-family: 'Vazirmatn', sans-serif;
        direction: rtl;
        background-color: #0F1115;
        color: #E6E8EC;
    }

    .stApp {
        background-color: #0F1115;
    }

    h1 {
        font-size: 2.4rem;
        font-weight: 700;
        text-align: center;
        margin-bottom: 5px;
    }

    .subtitle {
        text-align: center;
        color: #9AA0A6;
        margin-bottom: 40px;
    }

    .card {
        background-color: #181A20;
        padding: 25px;
        border-radius: 14px;
        border: 1px solid #2A2D36;
        margin-bottom: 25px;
    }

    .stButton > button {
        background-color: #5B8DEF;
        border-radius: 10px;
        height: 44px;
        font-weight: 600;
        border: none;
        transition: background 0.2s ease;
    }

    .stButton > button:hover {
        background-color: #4A7CDB;
    }

    [data-testid="stFileUploader"] {
        background-color: #181A20;
        border: 1px solid #2A2D36;
        border-radius: 12px;
        padding: 15px;
    }

    [data-testid="stDataFrame"] {
        border: 1px solid #2A2D36;
        border-radius: 12px;
    }

    </style>
    """, unsafe_allow_html=True)


# ==========================================================
# 🧠 DATA UTILITIES
# ==========================================================

def clean_currency(val):
    if pd.isna(val) or val == "":
        return 0.0
    return float(str(val).replace(",", "").replace("ریال", "").strip())


def find_header_row(df, keywords):
    for i in range(min(20, len(df))):
        if any(keyword in str(val) for val in df.iloc[i] for keyword in keywords):
            return i
    return 0


# ==========================================================
# 📊 PROCESSING LAYER
# ==========================================================

@st.cache_data
def process_sales(file):
    df_temp = pd.read_excel(file, header=None)
    header_row = find_header_row(df_temp, ["Date", "تاریخ"])
    df = pd.read_excel(file, header=header_row)

    df["Deposit"] = df.get("Deposit", 0).apply(clean_currency)
    df["Withdrawal"] = df.get("Withdrawal", 0).apply(clean_currency)

    result = df.groupby("Date", sort=False).agg({
        "Deposit": "sum",
        "Withdrawal": "sum"
    }).reset_index()

    result["Net"] = result["Deposit"] - result["Withdrawal"]

    return result


@st.cache_data
def process_statement(file):
    df_temp = pd.read_excel(file, header=None)
    header_row = find_header_row(df_temp, ["Date", "تاریخ"])
    df = pd.read_excel(file, header=header_row)

    df["Withdrawal"] = df.get("Withdrawal", 0).apply(clean_currency)

    result = df.groupby("Date", sort=False)["Withdrawal"].sum().reset_index()
    return result


# ==========================================================
# 📥 EXPORT SYSTEM
# ==========================================================

def export_excel(df, sheet_name="Report"):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    ws.append(list(df.columns))

    for row in dataframe_to_rows(df, index=False, header=False):
        ws.append(row)

    wb.save(output)
    return output.getvalue()


# ==========================================================
# 🖥 UI COMPONENTS
# ==========================================================

def render_header():
    st.markdown("<h1>داشبورد مالی کیمیا</h1>", unsafe_allow_html=True)
    st.markdown(
        "<div class='subtitle'>سیستم تحلیل حرفه‌ای تراکنش‌های بانکی</div>",
        unsafe_allow_html=True
    )


def render_sales_section():
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.subheader("📊 گزارش فروش")

    file = st.file_uploader("فایل اکسل فروش", type=["xlsx"], key="sales")

    if file:
        if st.button("پردازش فروش"):
            df = process_sales(file)
            st.success("پردازش با موفقیت انجام شد")
            st.dataframe(df, use_container_width=True)

            st.download_button(
                "دانلود خروجی اکسل",
                export_excel(df, "Sales"),
                f"Sales_{datetime.now().strftime('%Y%m%d')}.xlsx"
            )

    st.markdown("</div>", unsafe_allow_html=True)


def render_statement_section():
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.subheader("📈 آنالیز صورتحساب")

    file = st.file_uploader("فایل اکسل صورتحساب", type=["xlsx"], key="statement")

    if file:
        if st.button("پردازش صورتحساب"):
            df = process_statement(file)
            st.success("تحلیل انجام شد")
            st.dataframe(df, use_container_width=True)

            st.download_button(
                "دانلود خروجی اکسل",
                export_excel(df, "Statement"),
                f"Statement_{datetime.now().strftime('%Y%m%d')}.xlsx"
            )

    st.markdown("</div>", unsafe_allow_html=True)


# ==========================================================
# 🚀 MAIN APP
# ==========================================================

def main():
    st.set_page_config(
        page_title="Kimia Finance v6.0",
        page_icon="💎",
        layout="wide"
    )

    apply_theme()
    render_header()

    col1, col2 = st.columns(2)

    with col1:
        render_sales_section()

    with col2:
        render_statement_section()

    st.markdown(
        "<div style='text-align:center; color:#555; margin-top:40px;'>© 2026 Kimia Finance v6.0 Premium</div>",
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
