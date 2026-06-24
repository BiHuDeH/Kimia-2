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

def generate_styled_excel(df, sheet_name="Report"):
    """تولید خروجی اکسل با فرمت کاملاً یکسان و مشابه app (2).py"""
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

        df['Description'] = df['Description'].astype(str)

        # فیلترینگ مطابق با درخواست کاربر (استفاده از شرط AND بومی پانداس برای اطلس)
        card_to_card_mask = df['Description'].str.contains("انتقال از", na=False)
        fee_mask = df['Description'].str.contains("کارمزد", na=False)
        daily_withdrawal_mask = df['Description'].str.contains("انتقال وجه به سپرده 379.8000.10822179.1 به نام سپرده كوتاه مدت - مهران کلانتريان_سامانه بانکداری نوین", na=False)
        snap_deposit_mask = df['Description'].str.contains("غذا", na=False) & df['Description'].str.contains("اطلس", na=False)

        card_to_card_sum = df[card_to_card_mask].groupby('Date')['Deposit'].sum().reindex(unique_dates, fill_value=0)
        fee_sum = df[fee_mask].groupby('Date')['Withdrawal'].sum().reindex(unique_dates, fill_value=0)
        daily_withdrawal_sum = df[daily_withdrawal_mask].groupby('Date')['Withdrawal'].sum().reindex(unique_dates, fill_value=0)
        snap_deposit_sum = df[snap_deposit_mask].groupby('Date')['Deposit'].sum().reindex(unique_dates, fill_value=0)
        
        end_of_day_balance = df.sort_values(['Date', 'Time']).groupby('Date')['Balance'].last().reindex(unique_dates, fill_value=0)

        report = pd.DataFrame(index=unique_dates)
        report.index.name = 'Date'
        
        report['Card_to_Card'] = card_to_card_sum
        report['Fee'] = fee_sum
        report['Daily_Withdrawal'] = daily_withdrawal_sum
        report['Snap_Deposit'] = snap_deposit_sum
        report['End_of_Day_Balance'] = end_of_day_balance

        report['Sales'] = report['Card_to_Card'] / 1.1
        report['Tax'] = report['Card_to_Card'] - report['Sales']

        report = report.reset_index()

        # تغییر نام ستون‌ها بر اساس خواسته کاربر
        report.columns = ['تاریخ', 'کارت به کارت', 'کارمزد', 'برداشت روز', 'واریزی اسنپ', 'مانده آخر روز', 'فروش', 'مالیات']
        
        # مرتب‌سازی نهایی
        final_order = ['تاریخ', 'کارت به کارت', 'فروش', 'مالیات', 'کارمزد', 'برداشت روز', 'مانده آخر روز', 'واریزی اسنپ']
        
        return report[final_order], None

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
        if missing:
