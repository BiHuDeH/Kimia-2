import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

# --- تنظیمات صفحه ---
st.set_page_config(
    page_title="داشبورد مالی کیمیا",
    page_icon="💎",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- استایل‌دهی پیشرفته (تم تیره و مدرن) ---
def apply_modern_design():
    st.markdown("""
        <style>
            /* فراخوانی فونت وزیرمتن */
            @import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@100;300;400;500;700;900&display=swap');
            
            /* تنظیمات کلی بادی و فونت */
            html, body, [class*="css"] {
                font-family: 'Vazirmatn', sans-serif !important;
                direction: rtl;
                text-align: right;
            }
            
            /* پس‌زمینه اصلی */
            .stApp {
                background-color: #121212;
                color: #e0e0e0;
            }
            
            /* کارت‌های مدرن */
            .modern-card {
                background-color: #1E1E1E;
                padding: 25px;
                border-radius: 16px;
                border: 1px solid #333;
                box-shadow: 0 4px 20px rgba(0,0,0,0.4);
                margin-bottom: 25px;
                transition: transform 0.3s ease;
            }
            .modern-card:hover {
                border-color: #555;
            }

            /* استایل هدرها */
            h1, h2, h3 {
                color: #ffffff !important;
                font-weight: 800 !important;
                text-shadow: 0 2px 4px rgba(0,0,0,0.5);
            }
            
            /* دکمه‌ها */
            .stButton > button {
                width: 100%;
                border-radius: 12px;
                height: 50px;
                font-size: 16px;
                font-weight: 700;
                color: white;
                background: linear-gradient(90deg, #ff5500, #ff8800);
                border: none;
                box-shadow: 0 4px 15px rgba(255, 85, 0, 0.3);
                transition: all 0.3s ease;
            }
            /* دکمه دوم (آبی) */
            div[data-testid="stVerticalBlock"] > div:nth-child(4) .stButton > button {
                 background: linear-gradient(90deg, #00c6ff, #0072ff);
                 box-shadow: 0 4px 15px rgba(0, 114, 255, 0.3);
            }

            .stButton > button:hover {
                transform: translateY(-3px) scale(1.02);
                color: white !important;
            }

            /* باکس آپلود فایل */
            [data-testid="stFileUploader"] {
                background-color: #262626;
                border: 2px dashed #444;
                border-radius: 15px;
                padding: 20px;
            }
            [data-testid="stFileUploader"]:hover {
                border-color: #ff5500;
            }
            [data-testid="stFileUploader"] small {
                color: #888;
                font-family: 'Vazirmatn', sans-serif;
            }

            /* جداول */
            [data-testid="stDataFrame"] {
                border: 1px solid #333;
                border-radius: 10px;
                overflow: hidden;
            }
            
            /* پیام‌ها */
            .stSuccess, .stError, .stInfo {
                background-color: #1E1E1E !important;
                color: white !important;
                border: 1px solid #333;
                border-radius: 10px;
            }
            
            /* اسکرول بار */
            ::-webkit-scrollbar {
                width: 8px;
                height: 8px;
            }
            ::-webkit-scrollbar-track { background: #121212; }
            ::-webkit-scrollbar-thumb { background: #444; border-radius: 4px; }
            ::-webkit-scrollbar-thumb:hover { background: #666; }

        </style>
    """, unsafe_allow_html=True)

apply_modern_design()

# --- توابع کمکی ---

def to_persian_num(text):
    """تبدیل اعداد انگلیسی به فارسی"""
    if not isinstance(text, str):
        text = str(text)
    translation = text.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹")
    return text.translate(translation)

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
    """استانداردسازی نام ستون‌ها"""
    df.columns = df.columns.astype(str)
    rename_dict = {}
    used_columns = set()

    for std_name, aliases in col_map.items():
        for col in df.columns:
            if col in used_columns:
                continue
            if any(alias in col for alias in aliases):
                rename_dict[col] = std_name
                used_columns.add(col)
                break 
    
    df = df.rename(columns=rename_dict)
    df = df.loc[:, ~df.columns.duplicated()]
    return df

def generate_excel_download(df, sheet_name="Report", header_color="FF8800"):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.sheet_view.rightToLeft = True
    
    headers = list(df.columns)
    ws.append(headers)
    
    for cell in ws[1]:
        cell.font = Font(bold=True, name='Tahoma', color="FFFFFF")
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.fill = PatternFill(start_color=header_color, fill_type="solid")
    
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

# --- توابع پردازش ---

@st.cache_data(show_spinner=False)
def process_sales_data(file):
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
        
        # در بخش فروش، مانده معمولاً آخرین تراکنش است (منطق قبلی حفظ شد مگر اینکه اینجا هم نیاز به تغییر باشد)
        # فرض بر این است که تغییر خواسته شده فقط مربوط به بخش دوم (صورتحساب) است.
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
        for col in cols_order:
            if col not in final_df.columns:
                final_df[col] = 0
                
        return final_df[cols_order], None

    except Exception as e:
        return None, str(e)

@st.cache_data(show_spinner=False)
def process_statement_data(file):
    try:
        df_temp = pd.read_excel(file, header=None)
        header_row = find_header_row(df_temp, ["Date", "تاریخ"])
        df = pd.read_excel(file, header=header_row)
        
        col_map = {
            'Withdrawal': ['برداشت', 'بدهکار', 'Withdrawal', 'مبلغ برداشت'],
            'Balance': ['مانده', 'Balance', 'مانده حساب'],
            'Description': ['شرح', 'توضیحات', 'Description', 'شرح تراکنش'],
            'Date': ['تاریخ', 'Date', 'تاریخ تراکنش'],
            'Time': ['زمان', 'Time', 'زمان تراکنش']
        }
        
        df = standardize_columns(df, col_map)
        
        req_cols = ['Date', 'Description', 'Withdrawal', 'Balance']
        missing = [c for c in req_cols if c not in df.columns]
        if missing:
            return None, f"ستون‌های الزامی یافت نشدند: {', '.join(missing)}"

        df = df.dropna(subset=['Date'])
        for col in ['Withdrawal', 'Balance']:
            df[col] = df[col].apply(clean_currency)
        df['Description'] = df['Description'].fillna("").astype(str)

        # نکته مهم: طبق درخواست کاربر، مرتب‌سازی زمانی حذف شد تا ترتیب فایل اکسل حفظ شود.
        # "اولین سطر" در فایل اکسل برای هر تاریخ ملاک است.
        # if 'Time' in df.columns:
        #     df = df.sort_values(by=['Date', 'Time'])
        
        w_keywords = ["انتقال از", "برداشت از"]
        f_keywords = ["برداشت برای کارمزد", "دریافت کارمزد"]

        def calc_daily(group):
            # محاسبه جمع برداشت‌ها
            w_sum = group[
                group['Description'].apply(lambda x: any(k in x for k in w_keywords))
            ]['Withdrawal'].sum()

            # محاسبه جمع کارمزدها
            f_sum = group[
                group['Description'].apply(lambda x: any(k in x for k in f_keywords))
            ]['Withdrawal'].sum()

            # اصلاح منطق: مانده روز = اولین سطر موجود در فایل برای آن تاریخ
            # با فرض اینکه فایل بانکی نزولی (جدید به قدیم) است، اولین سطر معمولاً آخرین تراکنش است.
            first_bal = group['Balance'].iloc[0]

            return pd.Series({
                'برداشت روز': w_sum,
                'کارمزد': f_sum, # تغییر نام ستون طبق درخواست
                'مانده روز': first_bal
            })

        # نکته: groupby در پانداس ترتیب گروه ها را بر اساس کلید مرتب میکند، اما ترتیب سطرها داخل هر گروه 
        # (اگر sort=False نباشد) حفظ می‌شود.
        result_df = df.groupby('Date', sort=False).apply(calc_daily).reset_index()
        result_df = result_df.rename(columns={'Date': 'تاریخ'})
        
        return result_df, None

    except Exception as e:
        return None, str(e)

# --- رابط کاربری ---

def main():
    # هدر گرافیکی
    st.markdown("""
        <div style="text-align: center; padding: 20px;">
            <h1 style="background: -webkit-linear-gradient(right, #ff5500, #ffcc00); -webkit-background-clip: text; -webkit-text-fill-color: transparent; font-size: 3.5em; margin-bottom: 10px;">
                داشبورد مالی کیمیا
            </h1>
            <p style="color: #888; font-size: 1.2em;">سیستم هوشمند تحلیل و گزارش‌گیری تراکنش‌های بانکی</p>
        </div>
    """, unsafe_allow_html=True)

    # --- بخش اول: گزارش فروش ---
    st.markdown('<div class="modern-card">', unsafe_allow_html=True)
    
    st.markdown("### 📊 بخش اول: گزارش فروش")
    st.markdown("محاسبه فروش خالص، مالیات و واریزی‌ها.")
    
    upl_file_1 = st.file_uploader("انتخاب فایل اکسل فروش", type=["xlsx"], key="upl_1")
    
    if upl_file_1:
        if st.button("شروع پردازش فروش", key="btn_1"):
            with st.spinner("در حال تحلیل داده‌ها..."):
                res_df, err = process_sales_data(upl_file_1)
                
                if err:
                    st.error(f"خطا: {err}")
                else:
                    st.success("پردازش با موفقیت انجام شد")
                    
                    # نمایش فارسی
                    display_df = res_df.copy()
                    for col in display_df.columns:
                        if col != 'تاریخ':
                            display_df[col] = display_df[col].apply(lambda x: to_persian_num(f"{x:,.0f}"))
                    
                    st.dataframe(display_df, use_container_width=True)
                    
                    excel_data = generate_excel_download(res_df, "Sales Report", "FF8800")
                    st.download_button(
                        "📥 دانلود فایل اکسل نهایی",
                        excel_data,
                        f"Sales_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_1"
                    )
    st.markdown('</div>', unsafe_allow_html=True)

    # --- بخش دوم: گزارش صورتحساب ---
    st.markdown('<div class="modern-card" style="border-top: 4px solid #0072ff;">', unsafe_allow_html=True)
    
    st.markdown("### 📈 بخش دوم: آنالیز صورتحساب")
    st.markdown("تفکیک برداشت‌های روزانه و کارمزدها با محاسبه مانده از اولین سطر.")

    upl_file_2 = st.file_uploader("انتخاب فایل اکسل صورتحساب", type=["xlsx"], key="upl_2")
    
    if upl_file_2:
        if st.button("شروع پردازش صورتحساب", key="btn_2"):
            with st.spinner("در حال محاسبه مقادیر..."):
                res_df_2, err_2 = process_statement_data(upl_file_2)
                
                if err_2:
                    st.error(f"خطا: {err_2}")
                else:
                    st.success("تحلیل صورتحساب تکمیل شد")
                    
                    # نمایش فارسی
                    display_df_2 = res_df_2.copy()
                    for col in display_df_2.columns:
                        if col != 'تاریخ':
                            display_df_2[col] = display_df_2[col].apply(lambda x: to_persian_num(f"{x:,.0f}"))

                    st.dataframe(display_df_2, use_container_width=True)
                    
                    excel_data_2 = generate_excel_download(res_df_2, "Statement Analysis", "0072ff")
                    st.download_button(
                        "📥 دانلود فایل اکسل نهایی",
                        excel_data_2,
                        f"Statement_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_2"
                    )
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown("<div style='text-align: center; color: #555; font-size: 0.8rem; margin-top: 3rem;'>© 2026 Kimia Finance | v5.0</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
