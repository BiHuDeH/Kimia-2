import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

# --- پیکربندی صفحه ---
st.set_page_config(
    page_title="داشبورد مالی کیمیا",
    page_icon="💎",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- استایل‌دهی مدرن ---
def apply_custom_styles():
    st.markdown("""
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@100;300;400;500;700;900&display=swap');
            
            html, body, [class*="css"] {
                font-family: 'Vazirmatn', sans-serif;
                direction: rtl;
                text-align: right;
            }
            .stApp { background-color: #f3f4f6; }
            .card {
                background-color: white;
                padding: 2rem;
                border-radius: 1rem;
                box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
                margin-bottom: 2rem;
                border: 1px solid #e5e7eb;
            }
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
                transform: translateY(-1px);
            }
            [data-testid="stFileUploader"] {
                border: 2px dashed #cbd5e1;
                border-radius: 1rem;
                padding: 1.5rem;
                background-color: #f8fafc;
            }
            [data-testid="stDataFrame"] {
                border-radius: 0.75rem;
                border: 1px solid #e5e7eb;
                overflow: hidden;
            }
            h1, h2, h3 { color: #1e293b; font-weight: 800; }
        </style>
    """, unsafe_allow_html=True)

apply_custom_styles()

# --- توابع کمکی ---

def clean_currency(val):
    """تبدیل رشته‌های مالی به عدد"""
    if pd.isna(val) or val == "":
        return 0.0
    val_str = str(val).replace(",", "").replace("ریال", "").strip()
    try:
        return float(val_str)
    except ValueError:
        return 0.0

def find_header_row(df, keywords):
    """یافتن هوشمند سطر هدر"""
    for i in range(min(20, len(df))):
        row_values = df.iloc[i].astype(str).tolist()
        if any(keyword in val for val in row_values for keyword in keywords):
            return i
    return 0

def standardize_columns(df, col_map):
    """
    استانداردسازی نام ستون‌ها با جلوگیری از تکرار
    """
    df.columns = df.columns.astype(str)
    rename_dict = {}
    used_columns = set()

    # اولویت با نگاشت‌هایی است که تعریف کردیم
    for std_name, aliases in col_map.items():
        for col in df.columns:
            if col in used_columns:
                continue
            # بررسی تطابق نام ستون با الگوها
            if any(alias in col for alias in aliases):
                rename_dict[col] = std_name
                used_columns.add(col)
                break  # وقتی اولین ستون مناسب پیدا شد، می‌رویم سراغ بعدی تا تکراری نگیریم
    
    # تغییر نام
    df = df.rename(columns=rename_dict)
    
    # حذف ستون‌های تکراری احتمالی (اگر در فایل اصلی دو ستون دقیقاً یک نام داشتند)
    df = df.loc[:, ~df.columns.duplicated()]
    
    return df

def generate_excel_download(df, sheet_name="Report", header_color="DDEBF7"):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.sheet_view.rightToLeft = True
    
    headers = list(df.columns)
    ws.append(headers)
    
    for cell in ws[1]:
        cell.font = Font(bold=True, name='Vazirmatn')
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.fill = PatternFill(start_color=header_color, fill_type="solid")
    
    for r in dataframe_to_rows(df, index=False, header=False):
        ws.append(r)
            
    for r in ws.iter_rows(min_row=2):
        for c in r:
            if c.column_letter != 'A':
                c.number_format = '#,##0'
            c.alignment = Alignment(horizontal='center')
            c.font = Font(name='Vazirmatn')
            
    for i, col in enumerate(ws.columns, 1):
        ws.column_dimensions[get_column_letter(i)].width = 20
        
    wb.save(output)
    return output.getvalue()

# --- توابع پردازش ---

@st.cache_data(show_spinner=False)
def process_sales_data(file):
    """منطق محاسباتی بخش اول (فروش)"""
    try:
        df_temp = pd.read_excel(file, header=None)
        header_row = find_header_row(df_temp, ["Date", "تاریخ"])
        df = pd.read_excel(file, header=header_row)
        
        col_map = {
            'Deposit': ['واریز', 'بستانکار', 'Deposit', 'مبلغ واریز'],
            'Withdrawal': ['برداشت', 'بدهکار', 'Withdrawal', 'مبلغ برداشت'],
            'Balance': ['مانده', 'Balance', 'مانده حساب'],
            'Description': ['شرح', 'توضیحات', 'Description', 'شرح تراکنش'],
            'Date': ['تاریخ', 'Date', 'تاریخ تراکنش']
        }
        
        df = standardize_columns(df, col_map)
        
        if 'Date' not in df.columns:
            return None, "ستون تاریخ (Date) یافت نشد."

        df = df.dropna(subset=['Date'])
        for col in ['Deposit', 'Withdrawal', 'Balance']:
            if col in df.columns:
                df[col] = df[col].apply(clean_currency)
        
        df['Description'] = df['Description'].fillna("").astype(str)

        # فیلترها
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
        rep['balance'] = df.groupby('Date')['Balance'].last() if 'Balance' in df.columns else 0

        rep = rep.fillna(0)
        rep['sales'] = rep['card'] / 1.1
        rep['tax'] = rep['card'] - rep['sales']
        
        final_df = rep.reset_index().rename(columns={'index': 'تاریخ'})
        rename_map = {
            'Date': 'تاریخ', 'card': 'کارت به کارت', 'sales': 'فروش',
            'tax': 'مالیات', 'fee': 'کارمزد', 'withdraw': 'برداشت روز',
            'balance': 'مانده آخر روز', 'snap': 'واریزی اسنپ'
        }
        final_df = final_df.rename(columns=rename_map)
        
        cols_order = ['تاریخ', 'کارت به کارت', 'فروش', 'مالیات', 'کارمزد', 'برداشت روز', 'مانده آخر روز', 'واریزی اسنپ']
        # اگر ستونی موجود نبود، آن را با صفر پر کن تا ارور ندهد
        for col in cols_order:
            if col not in final_df.columns:
                final_df[col] = 0
                
        return final_df[cols_order], None

    except Exception as e:
        return None, str(e)

@st.cache_data(show_spinner=False)
def process_statement_data(file):
    """منطق محاسباتی بخش دوم (صورتحساب)"""
    try:
        df_temp = pd.read_excel(file, header=None)
        header_row = find_header_row(df_temp, ["Date", "تاریخ"])
        df = pd.read_excel(file, header=header_row)
        
        col_map = {
            'Withdrawal': ['برداشت', 'بدهکار', 'Withdrawal', 'مبلغ برداشت'],
            'Balance': ['مانده', 'Balance', 'مانده حساب'],
            'Description': ['شرح', 'توضیحات', 'Description', 'شرح تراکنش'],
            'Date': ['تاریخ', 'Date', 'تاریخ تراکنش'], # اولویت با تاریخ تراکنش است
            'Time': ['زمان', 'Time', 'زمان تراکنش']
        }
        
        # استفاده از تابع جدید برای جلوگیری از تکرار ستون‌ها
        df = standardize_columns(df, col_map)
        
        req_cols = ['Date', 'Description', 'Withdrawal', 'Balance']
        missing = [c for c in req_cols if c not in df.columns]
        if missing:
            return None, f"ستون‌های الزامی یافت نشدند: {', '.join(missing)}"

        df = df.dropna(subset=['Date'])
        for col in ['Withdrawal', 'Balance']:
            df[col] = df[col].apply(clean_currency)
        df['Description'] = df['Description'].fillna("").astype(str)

        if 'Time' in df.columns:
            df = df.sort_values(by=['Date', 'Time'])
        
        # منطق محاسباتی
        w_keywords = ["انتقال از", "برداشت از"]
        f_keywords = ["برداشت برای کارمزد", "دریافت کارمزد"]

        def calc_daily(group):
            w_sum = group[
                group['Description'].apply(lambda x: any(k in x for k in w_keywords))
            ]['Withdrawal'].sum()

            f_sum = group[
                group['Description'].apply(lambda x: any(k in x for k in f_keywords))
            ]['Withdrawal'].sum()

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

# --- رابط کاربری ---

def main():
    st.markdown("<h1 style='text-align: center; margin-bottom: 10px;'>داشبورد جامع مالی کیمیا</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; color: #64748b; margin-bottom: 40px;'>سیستم هوشمند تحلیل و گزارش‌گیری تراکنش‌های بانکی</p>", unsafe_allow_html=True)

    # --- بخش اول ---
    st.markdown('<div class="card">', unsafe_allow_html=True)
    col1_head, col1_img = st.columns([3, 1])
    with col1_head:
        st.header("بخش اول: گزارش فروش")
        st.markdown("محاسبه فروش خالص، مالیات و واریزی‌ها از فایل اکسل تراکنش‌ها.")
    with col1_img:
        st.image("https://cdn-icons-png.flaticon.com/512/3076/3076404.png", width=100)

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
                    
                    excel_data = generate_excel_download(res_df, "Sales Report", "DDEBF7")
                    st.download_button(
                        "📥 دانلود فایل اکسل",
                        excel_data,
                        f"Sales_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_1"
                    )
    st.markdown('</div>', unsafe_allow_html=True)

    # --- بخش دوم ---
    st.markdown('<div class="card">', unsafe_allow_html=True)
    col2_head, col2_img = st.columns([3, 1])
    with col2_head:
        st.header("بخش دوم: آنالیز صورتحساب")
        st.markdown("تفکیک برداشت‌های روزانه، کارمزدها و محاسبه مانده نهایی.")
    with col2_img:
        st.image("https://cdn-icons-png.flaticon.com/512/2344/2344143.png", width=100)

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
                    
                    excel_data_2 = generate_excel_download(res_df_2, "Statement Analysis", "E2EFDA")
                    st.download_button(
                        "📥 دانلود فایل اکسل",
                        excel_data_2,
                        f"Statement_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_2"
                    )
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown("<div style='text-align: center; color: #94a3b8; font-size: 0.8rem; margin-top: 3rem;'>© 2026 Kimia Finance | نسخه 3.1</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
