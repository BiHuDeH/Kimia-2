import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# --- تنظیمات صفحه ---
st.set_page_config(
    page_title="داشبورد مالی کیمیا",
    page_icon="💎",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- استایل‌دهی پیشرفته (UI/UX) ---
def apply_modern_design():
    st.markdown("""
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@100;300;400;500;700;900&display=swap');
            
            html, body, [class*="css"] {
                font-family: 'Vazirmatn', sans-serif !important;
                direction: rtl;
                text-align: right;
            }
            
            .stApp {
                background-color: #f8f9fa;
                color: #212529;
            }
            
            h1, h2, h3 {
                color: #0b3d91 !important;
                font-weight: 800 !important;
                text-shadow: 1px 1px 2px rgba(0,0,0,0.1);
            }

            .stTabs [data-baseweb="tab-list"] {
                gap: 10px;
                background-color: #e9ecef;
                padding: 10px 10px 0 10px;
                border-radius: 12px 12px 0 0;
            }
            .stTabs [data-baseweb="tab"] {
                height: 50px;
                background-color: white;
                border-radius: 8px 8px 0 0;
                padding: 10px 20px;
                font-weight: bold;
                color: #495057;
                border: 1px solid #dee2e6;
                border-bottom: none;
            }
            .stTabs [aria-selected="true"] {
                background-color: #0b3d91 !important;
                color: white !important;
            }

            [data-testid="stFileUploader"] {
                background-color: white;
                border: 2px dashed #adb5bd;
                border-radius: 15px;
                padding: 25px;
                transition: all 0.3s ease;
                box-shadow: 0 4px 6px rgba(0,0,0,0.05);
            }
            [data-testid="stFileUploader"]:hover {
                border-color: #0b3d91;
                background-color: #f1f5f9;
            }

            .stButton > button {
                width: 100%;
                border-radius: 12px;
                height: 50px;
                font-size: 16px;
                font-weight: 700;
                color: white;
                background: linear-gradient(135deg, #0b3d91, #1e90ff);
                border: none;
                box-shadow: 0 4px 15px rgba(11, 61, 145, 0.3);
                transition: transform 0.2s ease, box-shadow 0.2s ease;
            }
            .stButton > button:hover {
                transform: translateY(-2px);
                box-shadow: 0 6px 20px rgba(11, 61, 145, 0.4);
                color: white !important;
            }

            .stDownloadButton > button {
                background: linear-gradient(135deg, #28a745, #20c997) !important;
                box-shadow: 0 4px 15px rgba(40, 167, 69, 0.3) !important;
            }

            [data-testid="stDataFrame"] {
                border-radius: 12px;
                overflow: hidden;
                box-shadow: 0 4px 10px rgba(0,0,0,0.08);
            }
        </style>
    """, unsafe_allow_html=True)

apply_modern_design()

# --- توابع کمکی ---

def to_persian_num(text):
    if not isinstance(text, str): text = str(text)
    translation = text.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹")
    return text.translate(translation)

def clean_currency(val):
    if pd.isna(val) or val == "": return 0.0
    val_str = str(val).replace(",", "").replace("ریال", "").strip()
    try: return float(val_str)
    except ValueError: return 0.0

def find_header_row(df, keywords):
    for i in range(min(20, len(df))):
        row_values = df.iloc[i].astype(str).tolist()
        if any(str(keyword) in str(val) for val in row_values for keyword in keywords):
            return i
    return 0

def standardize_columns(df, col_map):
    df.columns = df.columns.astype(str)
    rename_dict = {}
    used_columns = set()
    for std_name, aliases in col_map.items():
        for col in df.columns:
            if col in used_columns: continue
            if any(str(alias) in str(col) for alias in aliases):
                rename_dict[col] = std_name
                used_columns.add(col)
                break 
    df = df.rename(columns=rename_dict)
    df = df.loc[:, ~df.columns.duplicated()]
    return df

def contains_all_words(text, phrase):
    """
    تابع جستجوی هوشمند با عملگر AND
    این تابع بررسی می‌کند که تمام واژه‌های موجود در یک عبارت، داخل متن (بدون در نظر گرفتن فاصله) وجود داشته باشند.
    """
    if pd.isna(text):
        return False
    text = str(text)
    words = phrase.split() # کلمات را بر اساس فاصله از هم جدا می‌کند
    return all(word in text for word in words)

def match_any_phrase(text, phrases_list):
    """بررسی اینکه آیا حداقل یکی از عبارات لیست، با شرط AND درون متن صدق می‌کند"""
    for phrase in phrases_list:
        if contains_all_words(text, phrase):
            return True
    return False

def generate_styled_excel(df, sheet_name="Report"):
    """تولید خروجی اکسل بر اساس استاندارد فرمت app2.py"""
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    ws.sheet_view.rightToLeft = True 

    # تنظیم استایل‌ها
    header_font = Font(bold=True, size=12, name='Tahoma')
    regular_font = Font(size=11, name='Tahoma')
    center_align = Alignment(horizontal="center", vertical="center")
    thin_border_side = Side(style="thin", color="333333")
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    
    headers = list(df.columns)
    ws.append(headers)
    
    # اعمال فرمت هدر
    for cell in ws[1]:
        cell.font = header_font
        cell.alignment = center_align
        cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        cell.border = thin_border
    
    # وارد کردن داده‌ها
    for row in df.itertuples(index=False, name=None):
        ws.append(row)

    # استایل‌دهی به سلول‌های بدنه
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.alignment = center_align
            cell.font = regular_font
            cell.border = thin_border
            if cell.column_letter != 'A':
                cell.number_format = '#,##0'

    # تنظیم عرض یکسان برای همه ستون‌ها
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 18

    # اعمال تیبل دیزاین اکسل
    if ws.max_row >= 2:
        last_col = get_column_letter(ws.max_column)
        tab = Table(displayName=f"Table_{sheet_name.replace(' ', '_')}", ref=f"A1:{last_col}{ws.max_row}")
        style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        tab.tableStyleInfo = style
        ws.add_table(tab)

    wb.save(output)
    return output.getvalue()

# --- پردازشگر گزارش پاسارگاد ---
@st.cache_data(show_spinner=False)
def process_pasargad(file):
    try:
        df = pd.read_excel(file, skiprows=2)
        
        expected_columns = [
            'Index', 'Branch Code', 'Branch', 'Date', 'Time', 
            'Document Number', 'Receipt Number', 'Check Number', 
            'Description', 'Withdrawal', 'Deposit', 'Balance', 'Notes'
        ]
        
        if len(df.columns) >= len(expected_columns):
            df.columns = expected_columns + list(df.columns[len(expected_columns):])
        else:
            return None, "ساختار فایل با الگوی استاندارد پاسارگاد مطابقت ندارد."

        if 'Date' not in df.columns: return None, "ستون Date یافت نشد."
        
        df = df.dropna(subset=['Date'])
        unique_dates = df['Date'].sort_values().unique()

        for col in ['Deposit', 'Withdrawal', 'Balance']:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

        # اعمال فیلترها با استفاده از تابع هوشمند AND
        card_mask = df['Description'].apply(lambda x: contains_all_words(x, "انتقال از"))
        fee_mask = df['Description'].apply(lambda x: contains_all_words(x, "کارمزد"))
        withdraw_mask = df['Description'].apply(lambda x: contains_all_words(x, "انتقال وجه به سپرده مهران کلانتریان"))
        snap_mask = df['Description'].apply(lambda x: contains_all_words(x, "غذا اطلس"))

        card_sum = df[card_mask].groupby('Date')['Deposit'].sum().reindex(unique_dates, fill_value=0)
        fee_sum = df[fee_mask].groupby('Date')['Withdrawal'].sum().reindex(unique_dates, fill_value=0)
        withdraw_sum = df[withdraw_mask].groupby('Date')['Withdrawal'].sum().reindex(unique_dates, fill_value=0)
        snap_sum = df[snap_mask].groupby('Date')['Deposit'].sum().reindex(unique_dates, fill_value=0)
        
        end_of_day_balance = df.sort_values(['Date', 'Time']).groupby('Date')['Balance'].last().reindex(unique_dates, fill_value=0)

        report = pd.DataFrame(index=unique_dates)
        report['card'] = card_sum
        report['fee'] = fee_sum
        report['withdraw'] = withdraw_sum
        report['snap'] = snap_sum
        report['balance'] = end_of_day_balance

        report['sales'] = report['card'] / 1.1
        report['tax'] = report['card'] - report['sales']

        report = report.reset_index().rename(columns={'index': 'تاریخ'})
        rename_map = {
            'Date': 'تاریخ', 'card': 'کارت به کارت', 'sales': 'فروش',
            'tax': 'مالیات', 'fee': 'کارمزد', 'withdraw': 'برداشت روز',
            'balance': 'مانده آخر روز', 'snap': 'واریزی اسنپ'
        }
        report = report.rename(columns=rename_map)
        
        cols_order = ['تاریخ', 'کارت به کارت', 'فروش', 'مالیات', 'کارمزد', 'برداشت روز', 'مانده آخر روز', 'واریزی اسنپ']
        return report[cols_order], None

    except Exception as e:
        return None, str(e)

# --- پردازشگر گزارش کارآفرین ---
@st.cache_data(show_spinner=False)
def process_karafrin(file):
    try:
        df_temp = pd.read_excel(file, header=None)
        header_row = find_header_row(df_temp, ["Date", "تاریخ"])
        df = pd.read_excel(file, header=header_row)
        
        col_map = {
            'Withdrawal': ['برداشت', 'بدهکار', 'Withdrawal', 'مبلغ برداشت'],
            'Balance': ['مانده', 'Balance', 'مانده حساب'],
            'Description': ['شرح', 'توضیحات', 'Description', 'شرح تراکنش'],
            'Date': ['تاریخ', 'Date', 'تاریخ تراکنش']
        }
        df = standardize_columns(df, col_map)
        
        req_cols = ['Date', 'Description', 'Withdrawal', 'Balance']
        missing = [c for c in req_cols if c not in df.columns]
        if missing: return None, f"ستون‌های الزامی یافت نشدند: {', '.join(missing)}"

        df = df.dropna(subset=['Date'])
        for col in ['Withdrawal', 'Balance']: df[col] = df[col].apply(clean_currency)
        
        w_keywords = ["انتقال از", "برداشت از"]
        f_keywords = ["برداشت برای کارمزد", "دریافت کارمزد"]

        def calc_daily(group):
            w_sum = group[group['Description'].apply(lambda x: match_any_phrase(x, w_keywords))]['Withdrawal'].sum()
            f_sum = group[group['Description'].apply(lambda x: match_any_phrase(x, f_keywords))]['Withdrawal'].sum()
            first_bal = group['Balance'].iloc[0]
            return pd.Series({'برداشت روز': w_sum, 'کارمزد': f_sum, 'مانده روز': first_bal})

        result_df = df.groupby('Date', sort=False).apply(calc_daily).reset_index()
        result_df = result_df.rename(columns={'Date': 'تاریخ'})
        return result_df, None

    except Exception as e:
        return None, str(e)

# --- رابط کاربری اصلی ---
def main():
    st.markdown("""
        <div style="text-align: center; padding: 10px 0 30px 0;">
            <h1 style="font-size: 3em; margin-bottom: 5px;">داشبورد مالی کیمیا</h1>
        </div>
    """, unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["🏦 گزارش مالی بانک پاسارگاد", "🏢 گزارش مالی بانک کارآفرین"])

    # ---------- تب پاسارگاد ----------
    with tab1:
        st.markdown("### 📑 گزارش مالی بانک پاسارگاد")
        
        upl_pasargad = st.file_uploader("فایل اکسل پاسارگاد را اینجا بکشید و رها کنید", type=["xlsx"], key="upl_pasargad")
        
        if upl_pasargad:
            if st.button("شروع پردازش پاسارگاد", key="btn_pasargad"):
                with st.spinner("در حال تحلیل فایل پاسارگاد..."):
                    res_pasargad, err_pasargad = process_pasargad(upl_pasargad)
                    
                    if err_pasargad:
                        st.error(f"خطا در پردازش: {err_pasargad}")
                    else:
                        st.success("گزارش پاسارگاد با موفقیت ایجاد شد!")
                        
                        disp_pasargad = res_pasargad.copy()
                        for col in disp_pasargad.columns:
                            if col != 'تاریخ': disp_pasargad[col] = disp_pasargad[col].apply(lambda x: to_persian_num(f"{x:,.0f}"))
                        
                        st.dataframe(disp_pasargad, use_container_width=True)
                        
                        excel_pasargad = generate_styled_excel(res_pasargad, "Pasargad Report")
                        st.download_button(
                            "📥 دانلود گزارش پاسارگاد",
                            excel_pasargad,
                            f"Pasargad_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_pasargad"
                        )

    # ---------- تب کارآفرین ----------
    with tab2:
        st.markdown("### 📑 گزارش مالی بانک کارآفرین")
        
        upl_karafrin = st.file_uploader("فایل اکسل کارآفرین را اینجا بکشید و رها کنید", type=["xlsx"], key="upl_karafrin")
        
        if upl_karafrin:
            if st.button("شروع پردازش کارآفرین", key="btn_karafrin"):
                with st.spinner("در حال تحلیل فایل کارآفرین..."):
                    res_karafrin, err_karafrin = process_karafrin(upl_karafrin)
                    
                    if err_karafrin:
                        st.error(f"خطا در پردازش: {err_karafrin}")
                    else:
                        st.success("گزارش کارآفرین با موفقیت ایجاد شد!")
                        
                        disp_karafrin = res_karafrin.copy()
                        for col in disp_karafrin.columns:
                            if col != 'تاریخ': disp_karafrin[col] = disp_karafrin[col].apply(lambda x: to_persian_num(f"{x:,.0f}"))
                        
                        st.dataframe(disp_karafrin, use_container_width=True)
                        
                        excel_karafrin = generate_styled_excel(res_karafrin, "Karafarin Report")
                        st.download_button(
                            "📥 دانلود گزارش کارآفرین",
                            excel_karafrin,
                            f"Karafarin_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="dl_karafrin"
                        )

if __name__ == "__main__":
    main()
