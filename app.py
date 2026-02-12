import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

# --- پیکربندی صفحه (باید اولین دستور باشد) ---
st.set_page_config(
    page_title="داشبورد مالی کیمیا",
    page_icon="💎",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- استایل‌دهی مدرن (CSS) ---
def apply_custom_styles():
    st.markdown("""
        <style>
            /* فونت وزیرمتن */
            @import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@100;300;400;500;700;900&display=swap');
            
            html, body, [class*="css"] {
                font-family: 'Vazirmatn', sans-serif;
                direction: rtl;
                text-align: right;
            }
            
            /* رنگ پس‌زمینه */
            .stApp {
                background-color: #f3f4f6;
            }
            
            /* کارت‌های سفید */
            .card {
                background-color: white;
                padding: 2rem;
                border-radius: 1rem;
                box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
                margin-bottom: 2rem;
                border: 1px solid #e5e7eb;
            }
            
            /* دکمه‌ها */
            .stButton > button {
                width: 100%;
                border-radius: 0.75rem;
                height: 3rem;
                font-weight: 700;
                background: linear-gradient(to right, #3b82f6, #2563eb);
                border: none;
                color: white;
                transition: all 0.3s;
            }
            .stButton > button:hover {
                background: linear-gradient(to right, #2563eb, #1d4ed8);
                box-shadow: 0 4px 12px rgba(37, 99, 235, 0.3);
                transform: translateY(-1px);
            }
            
            /* ناحیه آپلود */
            [data-testid="stFileUploader"] {
                border: 2px dashed #cbd5e1;
                border-radius: 1rem;
                padding: 1.5rem;
                background-color: #f8fafc;
            }
            
            /* جداول */
            [data-testid="stDataFrame"] {
                border-radius: 0.75rem;
                border: 1px solid #e5e7eb;
                overflow: hidden;
            }
            
            /* هدرها */
            h1, h2, h3 {
                color: #1e293b;
                font-weight: 800;
            }
            
            /* جداکننده */
            hr {
                border-color: #e5e7eb;
                margin: 2rem 0;
            }
        </style>
    """, unsafe_allow_html=True)

apply_custom_styles()

# --- توابع کمکی (Utility Functions) ---

def clean_currency(val):
    """تبدیل رشته‌های مالی (مثل '1,200,000 ریال') به عدد اعشاری"""
    if pd.isna(val) or val == "":
        return 0.0
    val_str = str(val).replace(",", "").replace("ریال", "").strip()
    try:
        return float(val_str)
    except ValueError:
        return 0.0

def find_header_row(df, keywords):
    """یافتن هوشمند سطر هدر بر اساس کلمات کلیدی"""
    # جستجو در 20 سطر اول
    for i in range(min(20, len(df))):
        row_values = df.iloc[i].astype(str).tolist()
        # اگر هر یک از کلمات کلیدی در این سطر پیدا شد
        if any(keyword in val for val in row_values for keyword in keywords):
            return i
    return 0

def generate_excel_download(df, sheet_name="Report", header_color="DDEBF7"):
    """تولید فایل اکسل با فرمت‌دهی زیبا"""
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.sheet_view.rightToLeft = True
    
    # نوشتن هدر
    headers = list(df.columns)
    ws.append(headers)
    
    # استایل هدر
    for cell in ws[1]:
        cell.font = Font(bold=True, name='Vazirmatn')
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.fill = PatternFill(start_color=header_color, fill_type="solid")
    
    # نوشتن داده‌ها
    for r in dataframe_to_rows(df, index=False, header=False):
        ws.append(r)
            
    # فرمت‌دهی اعداد (سه رقم سه رقم)
    for r in ws.iter_rows(min_row=2):
        for c in r:
            # فرض می‌کنیم ستون اول (تاریخ) متن است و بقیه عدد
            if c.column_letter != 'A':
                c.number_format = '#,##0'
            c.alignment = Alignment(horizontal='center')
            c.font = Font(name='Vazirmatn')
            
    # تنظیم عرض ستون‌ها
    for i, col in enumerate(ws.columns, 1):
        ws.column_dimensions[get_column_letter(i)].width = 20
        
    wb.save(output)
    return output.getvalue()

# --- توابع پردازش (Cached Logic) ---

@st.cache_data(show_spinner=False)
def process_sales_data(file):
    """منطق محاسباتی بخش اول (فروش)"""
    try:
        # خواندن فایل
        df_temp = pd.read_excel(file, header=None)
        header_row = find_header_row(df_temp, ["Date", "تاریخ"])
        df = pd.read_excel(file, header=header_row)
        
        # استانداردسازی ستون‌ها
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
            return None, "ستون تاریخ (Date) یافت نشد."

        # تمیزکاری داده‌ها
        df = df.dropna(subset=['Date'])
        for col in ['Deposit', 'Withdrawal', 'Balance']:
            if col in df.columns:
                df[col] = df[col].apply(clean_currency)
        df['Description'] = df['Description'].fillna("").astype(str)

        # فیلترها و گروه‌بندی (منطق ثابت قبلی)
        card_mask = df['Description'].str.contains("انتقال از", na=False)
        fee_mask = df['Description'].str.contains("کارمزد", na=False)
        withdraw_mask = df['Description'].str.contains("انتقال وجه", na=False)
        snap_mask = df['Description'].str.contains("مدرن سامانه غذارسان اطلس", na=False)

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
        rep['sales'] = rep['card'] / 1.1
        rep['tax'] = rep['card'] - rep['sales']
        
        # آماده‌سازی خروجی
        final_df = rep.reset_index().rename(columns={'index': 'تاریخ'})
        rename_map = {
            'Date': 'تاریخ', 'card': 'کارت به کارت', 'sales': 'فروش',
            'tax': 'مالیات', 'fee': 'کارمزد', 'withdraw': 'برداشت روز',
            'balance': 'مانده آخر روز', 'snap': 'واریزی اسنپ'
        }
        final_df = final_df.rename(columns=rename_map)
        cols_order = ['تاریخ', 'کارت به کارت', 'فروش', 'مالیات', 'کارمزد', 'برداشت روز', 'مانده آخر روز', 'واریزی اسنپ']
        return final_df[cols_order], None

    except Exception as e:
        return None, str(e)

@st.cache_data(show_spinner=False)
def process_statement_data(file):
    """منطق محاسباتی بخش دوم (صورتحساب)"""
    try:
        # خواندن فایل
        df_temp = pd.read_excel(file, header=None)
        header_row = find_header_row(df_temp, ["Date", "تاریخ"])
        df = pd.read_excel(file, header=header_row)
        
        # استانداردسازی ستون‌ها
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
        
        req_cols = ['Date', 'Description', 'Withdrawal', 'Balance']
        if any(c not in df.columns for c in req_cols):
            return None, f"ستون‌های الزامی یافت نشدند: {req_cols}"

        # تمیزکاری
        df = df.dropna(subset=['Date'])
        for col in ['Withdrawal', 'Balance']:
            df[col] = df[col].apply(clean_currency)
        df['Description'] = df['Description'].fillna("").astype(str)

        # مرتب‌سازی برای اطمینان از مانده صحیح (آخرین تراکنش روز)
        if 'Time' in df.columns:
            df = df.sort_values(by=['Date', 'Time'])
        
        # === منطق اصلی پردازش بخش دوم ===
        w_keywords = ["انتقال از", "برداشت از"]
        f_keywords = ["برداشت برای کارمزد", "دریافت کارمزد"]

        def calc_daily(group):
            # محاسبه جمع برداشت‌هایی که شامل کلمات کلیدی هستند
            w_sum = group[
                group['Description'].apply(lambda x: any(k in x for k in w_keywords))
            ]['Withdrawal'].sum()

            # محاسبه جمع کارمزدهایی که شامل کلمات کلیدی هستند
            f_sum = group[
                group['Description'].apply(lambda x: any(k in x for k in f_keywords))
            ]['Withdrawal'].sum()

            # آخرین مانده روز
            last_bal = group['Balance'].iloc[-1]

            return pd.Series({
                'برداشت روز': w_sum,
                'کارمزد روز': f_sum,
                'مانده روز': last_bal
            })

        result_df = df.groupby('Date').apply(calc_daily).reset_index()
        result_df = result_df.rename(columns={'Date': 'تاریخ'})
        
        return result_df, None

    except Exception as e:
        return None, str(e)

# --- رابط کاربری (UI) ---

def main():
    # هدر اصلی
    st.markdown("<h1 style='text-align: center; margin-bottom: 10px;'>داشبورد جامع مالی کیمیا</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #64748b; margin-bottom: 40px;'>سیستم هوشمند تحلیل و گزارش‌گیری تراکنش‌های بانکی</p>", unsafe_allow_html=True)

    # --- بخش اول: گزارش فروش ---
    st.markdown('<div class="card">', unsafe_allow_html=True)
    col1_head, col1_img = st.columns([3, 1])
    with col1_head:
        st.header("بخش اول: گزارش فروش")
        st.markdown("محاسبه فروش خالص، مالیات و واریزی‌ها از فایل اکسل تراکنش‌ها.")
    with col1_img:
        st.image("https://cdn-icons-png.flaticon.com/512/3076/3076404.png", width=100) # آیکون گزارش

    upl_file_1 = st.file_uploader("انتخاب فایل اکسل فروش", type=["xlsx"], key="upl_1")
    
    if upl_file_1:
        if st.button("📊 پردازش گزارش فروش", key="btn_1"):
            with st.spinner("در حال تحلیل داده‌ها..."):
                res_df, err = process_sales_data(upl_file_1)
                
                if err:
                    st.error(f"خطا: {err}")
                else:
                    st.success("پردازش با موفقیت انجام شد")
                    st.dataframe(res_df, use_container_width=True)
                    
                    excel_data = generate_excel_download(res_df, "Sales Report", "DDEBF7") # تم آبی
                    st.download_button(
                        "📥 دانلود فایل اکسل",
                        excel_data,
                        f"Sales_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_1"
                    )
    st.markdown('</div>', unsafe_allow_html=True)

    # --- بخش دوم: گزارش صورتحساب ---
    st.markdown('<div class="card">', unsafe_allow_html=True)
    col2_head, col2_img = st.columns([3, 1])
    with col2_head:
        st.header("بخش دوم: آنالیز صورتحساب")
        st.markdown("تفکیک برداشت‌های روزانه، کارمزدها و محاسبه مانده نهایی.")
    with col2_img:
        st.image("https://cdn-icons-png.flaticon.com/512/2344/2344143.png", width=100) # آیکون آنالیز

    upl_file_2 = st.file_uploader("انتخاب فایل اکسل صورتحساب", type=["xlsx"], key="upl_2")
    
    if upl_file_2:
        if st.button("📈 پردازش صورتحساب", key="btn_2"):
            with st.spinner("در حال محاسبه مقادیر..."):
                res_df_2, err_2 = process_statement_data(upl_file_2)
                
                if err_2:
                    st.error(f"خطا: {err_2}")
                else:
                    st.success("تحلیل صورتحساب تکمیل شد")
                    st.dataframe(res_df_2, use_container_width=True)
                    
                    excel_data_2 = generate_excel_download(res_df_2, "Statement Analysis", "E2EFDA") # تم سبز
                    st.download_button(
                        "📥 دانلود فایل اکسل",
                        excel_data_2,
                        f"Statement_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_2"
                    )
    st.markdown('</div>', unsafe_allow_html=True)
    
    # پاورقی
    st.markdown("<div style='text-align: center; color: #94a3b8; font-size: 0.8rem; margin-top: 3rem;'>© 2026 Kimia Finance | نسخه 3.0</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
