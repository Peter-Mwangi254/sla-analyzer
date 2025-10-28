# app.py
import streamlit as st
import pandas as pd
import io
from openpyxl.styles import Font, numbers
from openpyxl.drawing.image import Image
import base64
import matplotlib.pyplot as plt
import urllib.request
import tempfile
from PIL import Image as PILImage

# === PAGE CONFIG ===
st.set_page_config(page_title="CloudTech SLA Analyzer", layout="centered")

# === DARK MODE + STYLE ===
st.markdown("""
<style>
    .main { background: #0e1117; color: #fafafa; padding: 2rem; }
    .title { font-size: 2.8rem; font-weight: bold; color: #1E90FF; text-align: center; margin: 1rem 0; }
    .upload { border: 3px dashed #1E90FF; padding: 40px; text-align: center; border-radius: 12px; margin: 20px 0; }
    .upload:hover { background: #1a1f2e; }
    .metric { background: #1E90FF; color: white; padding: 15px; border-radius: 8px; text-align: center; font-size: 1.2em; margin: 10px 0; }
    .btn-download { background: #27ae60; color: white; padding: 12px 24px; border: none; border-radius: 6px; cursor: pointer; text-decoration: none; display: inline-block; }
    .footer { text-align: center; color: #888; margin-top: 50px; font-size: 0.9em; }
</style>
""", unsafe_allow_html=True)

# === HEADER ===
st.markdown("<div class='title'>CLOUDTECH SLA ANALYZER</div>", unsafe_allow_html=True)
st.markdown("<p style='text-align:center; color:#aaa;'>Upload any Excel → Get professional SLA report instantly</p>", unsafe_allow_html=True)

# === LOGO ===
logo_url = "https://i.imgur.com/5v5V5v5.png"  # Replace with your real logo
try:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
        urllib.request.urlretrieve(logo_url, tmp.name)
        logo = PILImage.open(tmp.name)
        st.image(logo, width=180, use_column_width=False)
except:
    st.markdown("<h2 style='text-align:center; color:#1E90FF;'>CLOUDTECH</h2>", unsafe_allow_html=True)

# === FILE UPLOAD ===
uploaded_file = st.file_uploader("Drop your Excel file here", type=["xlsx"], key="uploader")

if uploaded_file:
    with st.spinner("Analyzing your data..."):
        # === READ & AUTO-DETECT SHEET ===
        xl = pd.ExcelFile(uploaded_file)
        df = None
        valid_sheet = None
        for sheet in xl.sheet_names:
            temp_df = pd.read_excel(uploaded_file, sheet_name=sheet)
            cols = temp_df.columns.str.lower()
            if (cols.str.contains('number').any() and 
                cols.str.contains('created').any() and 
                cols.str.contains('actual.*end|work.*end|close', regex=True).any()):
                df = temp_df.copy()
                valid_sheet = sheet
                break

        if df is None:
            st.error("No valid sheet found. Need: Number, Created, Actual work end")
            st.stop()

        st.success(f"Found data in sheet: **{valid_sheet}** → {len(df)} rows")

        # === AUTO-DETECT COLUMNS ===
        cols_lower = df.columns.str.lower()
        num_col = df.columns[cols_lower.str.contains('number')][0]
        created_col = df.columns[cols_lower.str.contains('created')][0]
        end_col = df.columns[cols_lower.str.contains('actual.*end|work.*end|close', regex=True)][0]

        df = df[[num_col, created_col, end_col]].copy()
        df.columns = ['Number', 'Created', 'Actual work end']

        # === DATE CONVERSION ===
        def to_dt(x):
            try:
                if pd.notna(x) and isinstance(x, (int, float)) and x > 10000:
                    return pd.to_datetime(x, unit='D', origin='1899-12-30')
                return pd.to_datetime(x)
            except:
                return pd.NaT

        df['Created'] = df['Created'].apply(to_dt)
        df['Actual work end'] = df['Actual work end'].apply(to_dt)
        df = df.dropna(subset=['Created', 'Actual work end']).copy()

        if len(df) == 0:
            st.error("No valid date rows found!")
            st.stop()

        # === SLA CALCULATION ===
        df['hours'] = (df['Actual work end'] - df['Created']).dt.total_seconds() / 3600
        df['STATUS'] = df['hours'] > 24

        total = len(df)
        within = len(df[~df['STATUS']])
        past = len(df[df['STATUS']])
        avg_h = df['hours'].mean()
        compliance = 1 - (past / total) if total > 0 else 0
        past_pct = past / total if total > 0 else 0

        # === DISPLAY METRICS ===
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"<div class='metric'>Within SLA: <strong>{within}</strong></div>", unsafe_allow_html=True)
            st.markdown(f"<div class='metric'>Avg Close: <strong>{avg_h:.2f}h</strong></div>", unsafe_allow_html=True)
        with col2:
            st.markdown(f"<div class='metric'>Past SLA: <strong>{past}</strong></div>", unsafe_allow_html=True)
            st.markdown(f"<div class='metric'>Compliance: <strong>{compliance:.2%}</strong></div>", unsafe_allow_html=True)

        # === PIE CHART ===
        fig, ax = plt.subplots(figsize=(5, 3))
        ax.pie([within, past], labels=['Within SLA', 'Past SLA'], autopct='%1.2f%%', 
               colors=['#2ecc71', '#e74c3c'], startangle=90)
        ax.axis('equal')
        ax.set_title("SLA Compliance", fontsize=14, color='white')
        plt.tight_layout()
        buf = io.BytesIO()
        plt.savefig(buf, format='png', facecolor='#0e1117', edgecolor='none')
        buf.seek(0)
        chart_b64 = base64.b64encode(buf.read()).decode()
        st.image(f"data:image/png;base64,{chart_b64}", use_column_width=True)

        # === EXCEL OUTPUT ===
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df[['Number', 'Created', 'Actual work end', 'hours', 'STATUS']].to_excel(writer, 'Sheet1', index=False)

            summary = [
                ['', '', 'Ticket Within SLA', within, ''],
                ['', '', 'Closed past SLA', past, ''],
                ['', '', 'Avg Close (h)', round(avg_h, 2), ''],
                ['', '', 'Past SLA %', past_pct, ''],
                ['', '', 'SLA COMPLIANCE', compliance, '']
            ]
            pd.DataFrame(summary).to_excel(writer, 'Sheet1', index=False, header=False, startrow=len(df)+3)

            # === FORMAT EXCEL ===
            wb = writer.book
            ws = writer.sheets['Sheet1']
            bold = Font(bold=True)
            for cell in ws[1]: cell.font = bold
            for row in range(len(df)+4, len(df)+9):
                ws.cell(row=row, column=3).font = bold
            ws[f'D{len(df)+6}'].number_format = '0.00%'
            ws[f'D{len(df)+7}'].number_format = '0.00%'

        output.seek(0)
        excel_b64 = base64.b64encode(output.read()).decode()

        # === DOWNLOAD BUTTONS ===
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{excel_b64}" download="CloudTech_SLA_Report.xlsx" class="btn-download">Download Excel Report</a>', unsafe_allow_html=True)
        with col2:
            st.markdown("PDF coming soon...", unsafe_allow_html=True)

else:
    st.info("Upload your Excel file to generate SLA report")

# === FOOTER ===
st.markdown("<div class='footer'>© 2025 CloudTech | SLA Analyzer v2.0</div>", unsafe_allow_html=True)