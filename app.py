import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# --- Page Configuration ---
st.set_page_config(
    page_title="داشبورد مالی کیمیا",
    page_icon="💎",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- Custom CSS (Modern, Minimal, Light Theme, Vazirmatn Font, RTL) ---
def set_custom_style():
    st.markdown(
        """
        <style>
            /* Import Vazirmatn Font */
            @import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@100;300;400;500;700;900&display=swap');

            /* Global Styles */
            html, body, [class*="css"] {
                font-family: 'Vazirmatn', sans-serif;
                direction: rtl;
                text-align: right;
            }

            /* Main Background */
            .stApp {
                background-color: #F3F4F6; /* Light gray background */
                color: #1F2937; /* Dark gray text for readability */
            }

            /* Headings */
            h1, h2, h3 {
                color: #111827; /* Almost black for headings */
                font-weight: 700;
            }

            /* Report Container Box */
            .report-box {
                background-color: #FFFFFF;
                padding: 30px;
                border-radius: 16px;
                box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
                margin-bottom: 30px;
                border: 1px solid #E5E7EB;
            }

            /* Buttons */
            div.stButton > button {
                background-color: #3B82F6; /* Modern Blue */
                color: white;
                border-radius: 12px;
                padding: 12px 24px;
                font-weight: 600;
                border: none;
                box-shadow: 0 2px 4px rgba(59, 130, 246, 0.2);
                transition: all 0.2s ease-in-out;
                width: 100%;
            }
            div.stButton > button:hover {
                background-color: #2563EB; /* Darker blue on hover */
                box-shadow: 0 4px 8px rgba(59, 130, 246, 0.3);
                transform: translateY(-1px);
            }
            div.stButton > button:active {
                 transform: translateY(0px);
            }

            /* File Uploader Area */
            div[data-testid="stFileUploader"] {
                border: 2px dashed #D1D5DB;
                background-color: #F9FAFB;
                border-radius: 12px;
                padding: 20px;
                text-align: center;
            }
            div[data-testid="stFileUploader"] section {
                 background-color: transparent;
            }

            /* Dataframes */
            div[data-testid="stDataFrame"] {
                direction: rtl;
                border-radius: 12px;
                overflow: hidden;
                border: 1px solid #E5E7EB;
            }

            /* Success/Error Messages */
            .stSuccess, .stError, .stWarning, .stInfo {
                direction: rtl;
                text-align: right;
                border-radius: 10px;
            }
            
            /* Divider */
            hr {
                margin-top: 2rem;
                margin-bottom: 2rem;
                border: 0;
                border-top: 1px solid #E5E7EB;
            }
            
            /* Images alignment */
            div[data-testid="stImage"] {
                display: flex;
                justify-content: center;
                margin-bottom: 20px;
            }
            div[data-testid="stImage"] img {
                max-height: 250px;
                object-fit: contain;
            }

        </style>
        """,
        unsafe_allow_html=True
    )

set_custom_style()

# --- Helper Functions (Shared) ---

def to_float(val):
    """Clean and convert monetary strings to float."""
    if pd.isna(val) or val == "":
        return 0.0
    val_str = str(val).replace(",", "").replace("ریال", "").strip()
    try:
        return float(val_str)
    except ValueError:
        return 0.0

def find_header_row(df, keywords):
    """Dynamically find the header row based on keywords."""
    for i in range(min(20, len(df))):
        row_values = df.iloc[i].astype(str).tolist()
        if any(keyword in val for val in row_values for keyword in keywords):
            return i
    return 0

def dataframe_to_rows(df, index=False, header=False):
    """Helper to convert DF to rows for openpyxl"""
    from openpyxl.utils.dataframe import dataframe_to_rows as dtr
    return dtr(df, index=index, header=header)

# ==========================================
# بخش اول: گزارش مالی فروش (کد قبلی)
# ==========================================

def process_section_1():
    st.markdown("<div class='report-box'>", unsafe_allow_html=True)
    
    col1, col2 = st.columns([3, 1])
    with col1:
        st.header("بخش اول: گزارش تجمیعی فروش")
        st.markdown("محاسبه فروش، مالیات و واریزی‌ها بر اساس فایل تراکنش‌ها.")
    with col2:
        # استفاده از یک تصویر ایلاستریشن مدرن برای بخش فروش
        st.image("https://cdn.dribbble.com/users/77121/screenshots/15677677/media/4914b76026533f283d4c07040c0723db.png?resize=400x300&vertical=center", use_container_width=True)

    
    uploaded_file = st.file_uploader("فایل اکسل فروش را بارگذاری کنید", type=["xlsx"], key="uploader_1")

    if uploaded_file is not None:
        if st.button("📊 پردازش فایل فروش", key="btn_process_1"):
            with st.spinner('در حال پردازش داده‌ها...'):
                try:
                    # 1. Load Data
                    try:
                        df_temp = pd.read_excel(uploaded_file, header=None)
                        header_row = find_header_row(df_temp, ["Date", "تاریخ"])
                        df = pd.read_excel(uploaded_file, header=header_row)
                    except Exception as e:
                        st.error(f"خطا در خواندن فایل: {e}")
                        return

                    # 2. Standardize Columns
                    col_map = {
                        'Deposit': ['واریز', 'بستانکار', 'Deposit', 'مبلغ واریز'],
                        'Withdrawal': ['برداشت', 'بدهکار', 'Withdrawal', 'مبلغ برداشت'],
                        'Balance': ['مانده', 'Balance', 'مانده حساب'],
                        'Description': ['شرح', 'توضیحات', 'Description', 'شرح تراکنش'],
                        'Date': ['تاریخ', 'Date', 'تاریخ تراکنش']
                    }
                    
                    for col in df.columns:
                        for std, aliases in col_map.items():
                            if any(alias in str(col) for alias in aliases):
                                df.rename(columns={col: std}, inplace=True)
                    
                    if 'Date' not in df.columns:
                        st.error("ستون تاریخ (Date) یافت نشد.")
                        return

                    # 3. Clean Data
                    df = df.dropna(subset=['Date'])
                    for col in ['Deposit', 'Withdrawal', 'Balance']:
                        if col in df.columns:
                            df[col] = df[col].apply(to_float)
                    
                    df['Description'] = df['Description'].fillna("").astype(str)

                    # 4. Filters (Old Logic)
                    card_mask = df['Description'].str.contains("انتقال از", na=False)
                    fee_mask = df['Description'].str.contains("کارمزد", na=False)
                    withdraw_mask = df['Description'].str.contains("انتقال وجه", na=False)
                    snap_mask = df['Description'].str.contains("مدرن سامانه غذارسان اطلس", na=False)

                    # 5. Grouping
                    dates = df['Date'].unique()
                    rep = pd.DataFrame(index=dates)
                    
                    rep['card'] = df[card_mask].groupby('Date')['Deposit'].sum()
                    rep['fee'] = df[fee_mask].groupby('Date')['Withdrawal'].sum()
                    rep['withdraw'] = df[withdraw_mask].groupby('Date')['Withdrawal'].sum()
                    rep['snap'] = df[snap_mask].groupby('Date')['Deposit'].sum()
                    
                    if 'Balance' in df.columns:
                        rep['balance'] = df.groupby('Date')['Balance'].last()
                    else:
                        rep['balance'] = 0

                    rep = rep.fillna(0)

                    # 6. Calculations
                    rep['sales'] = rep['card'] / 1.1
                    rep['tax'] = rep['card'] - rep['sales']
                    
                    # Format for output
                    final_df = rep.reset_index().rename(columns={'index': 'تاریخ'})
                    output_columns = {
                        'Date': 'تاریخ',
                        'card': 'کارت به کارت',
                        'sales': 'فروش',
                        'tax': 'مالیات',
                        'fee': 'کارمزد',
                        'withdraw': 'برداشت روز',
                        'balance': 'مانده آخر روز',
                        'snap': 'واریزی اسنپ'
                    }
                    final_df = final_df.rename(columns=output_columns)
                    
                    cols_to_keep = ['تاریخ', 'کارت به کارت', 'فروش', 'مالیات', 'کارمزد', 'برداشت روز', 'مانده آخر روز', 'واریزی اسنپ']
                    final_df = final_df[cols_to_keep]

                    st.success("✅ پردازش بخش اول با موفقیت انجام شد.")
                    st.dataframe(final_df.head(), use_container_width=True)

                    # Download Button
                    excel_data = generate_excel_section1(final_df)
                    st.download_button(
                        label="📥 دانلود گزارش فروش (Excel)",
                        data=excel_data,
                        file_name=f"Sales_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_btn_1"
                    )

                except Exception as e:
                    st.error(f"❌ خطا در پردازش بخش اول: {e}")
    st.markdown("</div>", unsafe_allow_html=True)


# ==========================================
# بخش دوم: گزارش برداشت و کارمزد (منطق جدید)
# ==========================================

def process_section_2():
    st.markdown("<div class='report-box' style='margin-top: 40px;'>", unsafe_allow_html=True)
    
    col1, col2 = st.columns([3, 1])
    with col1:
        st.header("بخش دوم: آنالیز برداشت و کارمزد")
        st.markdown("تفکیک برداشت‌های روزانه و کارمزدها بر اساس شرح تراکنش.")
    with col2:
        # استفاده از یک تصویر ایلاستریشن مدرن برای بخش تحلیل داده
        st.image("https://cdn.dribbble.com/users/77121/screenshots/11976029/media/906d0774601b27d3464340d4eb027d1c.png?resize=400x300&vertical=center", use_container_width=True)

    uploaded_file = st.file_uploader("فایل اکسل صورتحساب را بارگذاری کنید", type=["xlsx"], key="uploader_2")

    if uploaded_file is not None:
        if st.button("📈 پردازش صورتحساب", key="btn_process_2"):
             with st.spinner('در حال تحلیل صورتحساب...'):
                try:
                    # 1. Load Data
                    try:
                        df_temp = pd.read_excel(uploaded_file, header=None)
                        header_row = find_header_row(df_temp, ["Date", "تاریخ"])
                        df = pd.read_excel(uploaded_file, header=header_row)
                    except Exception as e:
                        st.error(f"خطا در خواندن فایل: {e}")
                        return

                    # 2. Standardize Columns
                    col_map = {
                        'Withdrawal': ['برداشت', 'بدهکار', 'Withdrawal', 'مبلغ برداشت'],
                        'Balance': ['مانده', 'Balance', 'مانده حساب'],
                        'Description': ['شرح', 'توضیحات', 'Description', 'شرح تراکنش'],
                        'Date': ['تاریخ', 'Date', 'تاریخ تراکنش'],
                        'Time': ['زمان', 'Time', 'زمان تراکنش']
                    }
                    
                    for col in df.columns:
                        for std, aliases in col_map.items():
                            if any(alias in str(col) for alias in aliases):
                                df.rename(columns={col: std}, inplace=True)
                    
                    # Validation
                    required_cols = ['Date', 'Description', 'Withdrawal', 'Balance']
                    missing = [c for c in required_cols if c not in df.columns]
                    if missing:
                        st.error(f"ستون‌های الزامی زیر یافت نشدند: {missing}")
                        return

                    # 3. Clean Data
                    df = df.dropna(subset=['Date'])
                    for col in ['Withdrawal', 'Balance']:
                        df[col] = df[col].apply(to_float)
                    
                    df['Description'] = df['Description'].fillna("").astype(str)

                    if 'Time' in df.columns:
                        df = df.sort_values(by=['Date', 'Time'])
                    else:
                        pass

                    # 4. New Logic Implementation
                    w_keywords = ["انتقال از", "برداشت از"]
                    f_keywords = ["برداشت برای کارمزد", "دریافت کارمزد"]

                    def calculate_daily_stats(group):
                        w_sum = group[
                            group['Description'].apply(lambda x: any(k in x for k in w_keywords))
                        ]['Withdrawal'].sum()

                        f_sum = group[
                            group['Description'].apply(lambda x: any(k in x for k in f_keywords))
                        ]['Withdrawal'].sum()

                        last_balance = group['Balance'].iloc[-1]

                        return pd.Series({
                            'برداشت روز': w_sum,
                            'کارمزد روز': f_sum,
                            'مانده روز': last_balance
                        })

                    result_df = df.groupby('Date').apply(calculate_daily_stats).reset_index()

                    # 5. Display and Download
                    st.success("✅ پردازش بخش دوم با موفقیت انجام شد.")
                    
                    result_df = result_df.rename(columns={'Date': 'تاریخ'})
                    
                    st.dataframe(result_df, use_container_width=True)

                    # Generate Excel
                    excel_data = generate_excel_section2(result_df)
                    st.download_button(
                        label="📥 دانلود گزارش صورتحساب (Excel)",
                        data=excel_data,
                        file_name=f"Withdrawal_Fee_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_btn_2"
                    )

                except Exception as e:
                    st.error(f"❌ خطا در پردازش بخش دوم: {e}")
    st.markdown("</div>", unsafe_allow_html=True)


# --- Excel Generation Functions ---

def generate_excel_section1(df):
    """Generates Excel for Section 1"""
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.sheet_view.rightToLeft = True
    
    headers = list(df.columns)
    ws.append(headers)
    
    for cell in ws[1]:
        cell.font = Font(bold=True, name='Tahoma')
        cell.alignment = Alignment(horizontal='center')
        cell.fill = PatternFill(start_color="DDEBF7", fill_type="solid")
    
    for r in dataframe_to_rows(df, index=False, header=False):
        ws.append(r)
            
    for r in ws.iter_rows(min_row=2):
        for c in r:
            if c.column_letter != 'A':
                c.number_format = '#,##0'
            c.alignment = Alignment(horizontal='center')
            
    for i, col in enumerate(ws.columns, 1):
        ws.column_dimensions[get_column_letter(i)].width = 18
        
    wb.save(output)
    return output.getvalue()

def generate_excel_section2(df):
    """Generates Excel for Section 2"""
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.sheet_view.rightToLeft = True
    
    headers = list(df.columns)
    ws.append(headers)
    
    for cell in ws[1]:
        cell.font = Font(bold=True, name='Tahoma')
        cell.alignment = Alignment(horizontal='center')
        cell.fill = PatternFill(start_color="E2EFDA", fill_type="solid")
    
    for r in dataframe_to_rows(df, index=False, header=False):
        ws.append(r)
            
    for r in ws.iter_rows(min_row=2):
        for c in r:
            if c.column_letter != 'A':
                c.number_format = '#,##0'
            c.alignment = Alignment(horizontal='center')
            
    for i, col in enumerate(ws.columns, 1):
        ws.column_dimensions[get_column_letter(i)].width = 20
        
    wb.save(output)
    return output.getvalue()

# --- Main Application ---
def main():
    st.title("داشبورد جامع مالی کیمیا")
    st.markdown("نسخه پایدار - طراحی مینیمال")
    st.markdown("---")
    
    # Section 1
    process_section_1()
    
    # Section 2
    process_section_2()
    
    st.markdown("<div style='text-align: center; color: #6B7280; margin-top: 50px;'>© 2024 Kimia Finance Dashboard</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
