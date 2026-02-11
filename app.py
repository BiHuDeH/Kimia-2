import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime  # <--- این خط جا افتاده بود
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# Version and Update Information
SCRIPT_VERSION = "v2.1 (Fixed)"
UPDATE_DATE = "2026-02-03"

# --- Page Configuration ---
st.set_page_config(page_title="Financial Data Report", page_icon="📊", layout="wide")

# Custom CSS for UI enhancements
def set_custom_style():
    st.markdown(
        f"""
        <style>
            .stApp {{
                background-color: #FAF3E0;
            }}
            h1, h2, h3 {{
                color: #003366;
            }}
            .stButton > button {{
                background-color: #003366;
                color: white;
                border-radius: 8px;
                width: 100%;
            }}
            .stButton > button:hover {{
                background-color: #00509E;
            }}
            /* Info box style */
            .info-box {{
                font-size: small; 
                text-align: right; 
                color: #666; 
                padding: 10px;
                border-top: 1px solid #ccc;
                margin-top: 20px;
            }}
        </style>
        """,
        unsafe_allow_html=True
    )

# --- Logic Functions ---

def create_styled_report(df):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Report"
    ws.sheet_view.rightToLeft = True  # Right-to-left for Persian

    # Styles
    header_font = Font(bold=True, size=12, name='Tahoma')
    regular_font = Font(size=11, name='Tahoma')
    center_align = Alignment(horizontal="center", vertical="center")
    thin_border_side = Side(style="thin", color="333333")
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    # Columns
    ordered_columns = ['تاریخ', 'کارت به کارت', 'فروش', 'مالیات', 'کارمزد', 'برداشت روز', 'مانده آخر روز', 'واریزی اسنپ']
    
    # Ensure all columns exist
    for col in ordered_columns:
        if col not in df.columns:
            df[col] = 0

    df_export = df[ordered_columns].copy()

    # Write Header
    ws.append(ordered_columns)
    for cell in ws[1]:
        cell.font = header_font
        cell.alignment = center_align
        cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        cell.border = thin_border

    # Write Data
    for row in df_export.itertuples(index=False, name=None):
        ws.append(row)

    # Apply Styles & Formats
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.font = regular_font
            cell.alignment = center_align
            cell.border = thin_border
            
            # Format Numbers (Skip Date column which is usually first)
            if cell.column_letter != 'A':
                cell.number_format = '#,##0'  # Format as 1,000

    # Auto-fit columns
    for col in ws.columns:
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = 18

    # Add Table Feature
    if ws.max_row >= 2:
        last_col = get_column_letter(ws.max_column)
        tab = Table(displayName="ReportTable", ref=f"A1:{last_col}{ws.max_row}")
        style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        tab.tableStyleInfo = style
        ws.add_table(tab)

    wb.save(output)
    output.seek(0)
    return output.getvalue()

@st.cache_data(show_spinner=False)
def process_data(df):
    # Ensure Date column exists and drop empty ones
    if 'Date' not in df.columns:
        return pd.DataFrame() 
        
    df = df.dropna(subset=['Date'])
    
    unique_dates = df['Date'].sort_values().unique()

    # Convert numeric columns to float just in case
    cols_to_numeric = ['Deposit', 'Withdrawal', 'Balance']
    for col in cols_to_numeric:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # Filters
    df['Description'] = df['Description'].astype(str)
    
    card_to_card_mask = df['Description'].str.contains("انتقال از", na=False)
    fee_mask = df['Description'].str.contains("کارمزد", na=False)
    daily_withdrawal_mask = df['Description'].str.contains("انتقال وجه", na=False)
    snap_deposit_mask = df['Description'].str.contains("مدرن سامانه غذارسان اطلس", na=False)

    # Grouping
    grouped = df.groupby('Date')
    
    card_to_card_sum = df[card_to_card_mask].groupby('Date')['Deposit'].sum().reindex(unique_dates, fill_value=0)
    fee_sum = df[fee_mask].groupby('Date')['Withdrawal'].sum().reindex(unique_dates, fill_value=0)
    daily_withdrawal_sum = df[daily_withdrawal_mask].groupby('Date')['Withdrawal'].sum().reindex(unique_dates, fill_value=0)
    snap_deposit_sum = df[snap_deposit_mask].groupby('Date')['Deposit'].sum().reindex(unique_dates, fill_value=0)
    
    # Logic for End of Day Balance
    end_of_day_balance = df.sort_values(['Date', 'Time']).groupby('Date')['Balance'].last().reindex(unique_dates, fill_value=0)

    # Create Report DataFrame
    report = pd.DataFrame(index=unique_dates)
    report.index.name = 'Date'
    
    report['Card_to_Card'] = card_to_card_sum
    report['Fee'] = fee_sum
    report['Daily_Withdrawal'] = daily_withdrawal_sum
    report['Snap_Deposit'] = snap_deposit_sum
    report['End_of_Day_Balance'] = end_of_day_balance

    # Calculations
    report['Sales'] = report['Card_to_Card'] / 1.1
    report['Tax'] = report['Card_to_Card'] - report['Sales']

    # Reset index
    report = report.reset_index()

    # Rename to Persian
    report.columns = ['تاریخ', 'کارت به کارت', 'کارمزد', 'برداشت روز', 'واریزی اسنپ', 'مانده آخر روز', 'فروش', 'مالیات']
    
    # Reorder columns
    final_order = ['تاریخ', 'کارت به کارت', 'فروش', 'مالیات', 'کارمزد', 'برداشت روز', 'مانده آخر روز', 'واریزی اسنپ']
    return report[final_order]

# --- Main App ---
def main():
    set_custom_style()
    
    with st.sidebar:
        st.header("تنظیمات")
        st.info(f"نسخه برنامه: {SCRIPT_VERSION}\n\nتاریخ به‌روزرسانی: {UPDATE_DATE}")

    st.title("📊 گزارش‌گیری مالی")
    st.write("لطفاً فایل اکسل گردش حساب را بارگذاری کنید.")

    uploaded_file = st.file_uploader("انتخاب فایل (Excel)", type=["xlsx"])
    
    if uploaded_file:
        try:
            # Read Data
            if uploaded_file.type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
                df = pd.read_excel(uploaded_file, skiprows=2)
                
                # Standardize Columns
                expected_columns = [
                    'Index', 'Branch Code', 'Branch', 'Date', 'Time', 
                    'Document Number', 'Receipt Number', 'Check Number', 
                    'Description', 'Withdrawal', 'Deposit', 'Balance', 'Notes'
                ]
                
                if len(df.columns) >= len(expected_columns):
                    df.columns = expected_columns + list(df.columns[len(expected_columns):])
                else:
                    st.error("ساختار فایل اکسل با الگوی استاندارد مطابقت ندارد.")
                    return
            else:
                st.error("لطفا فقط فایل اکسل بارگذاری کنید.")
                return

            # Process Data
            with st.spinner('در حال پردازش اطلاعات...'):
                report_df = process_data(df)

            if not report_df.empty:
                st.success("پردازش با موفقیت انجام شد!")
                
                tab1, tab2 = st.tabs(["📋 پیش‌نمایش جدول", "📥 دانلود گزارش"])

                with tab1:
                    st.dataframe(
                        report_df,
                        use_container_width=True,
                        column_config={
                            "تاریخ": st.column_config.TextColumn("تاریخ"),
                            "کارت به کارت": st.column_config.NumberColumn(format="%.0f"),
                            "فروش": st.column_config.NumberColumn(format="%.0f"),
                            "مالیات": st.column_config.NumberColumn(format="%.0f"),
                            "کارمزد": st.column_config.NumberColumn(format="%.0f"),
                            "برداشت روز": st.column_config.NumberColumn(format="%.0f"),
                            "مانده آخر روز": st.column_config.NumberColumn(format="%.0f"),
                            "واریزی اسنپ": st.column_config.NumberColumn(format="%.0f"),
                        }
                    )

                with tab2:
                    excel_data = create_styled_report(report_df)
                    st.download_button(
                        label="📥 دانلود فایل اکسل نهایی",
                        data=excel_data,
                        file_name=f"Financial_Report_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.warning("داده‌ای برای نمایش یافت نشد.")

        except Exception as e:
            st.error(f"خطایی رخ داد: {e}")

if __name__ == "__main__":
    main()
