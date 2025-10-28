# app.py
import streamlit as st
import pandas as pd
import io
import base64
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.drawing.image import Image as XLImage
import tempfile
import urllib.request
from PIL import Image as PILImage

st.set_page_config(page_title="CloudTech SLA", layout="centered")

# === STYLE ===
st.markdown("""
<style>
    .main { background: #0a0e17; color: #fafafa; padding: 2rem; }
    .title { font-size: 2.8rem; font-weight: bold; color: #1E90FF; text-align: center; }
    .upload { border: 3px dashed #1E90FF; padding: 40px; border-radius: 12px; text-align: center; }
    .metric { background: #1E90FF; color: white; padding: 15px; border-radius: 8px; text-align: center; font-size: 1.2em; margin: 10px 0; }
    .btn { background: #27ae60; color: white; padding: 12px 24px; border: none; border-radius: 6px; cursor: pointer; text-decoration: none; }
    .footer { text-align: center; color: #888; margin-top: 50px; }
</style>
""", unsafe_allow_html=True)

st.markdown("<div class='title'>CLOUDTECH SLA ANALYZER</div>", unsafe_allow_html=True)

# === LOGO ===
logo_url = "https://i.imgur.com/5v5V5v5.png"
try:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as f:
        urllib.request.urlretrieve(logo_url, f.name)
        st.image(f.name, width=180)
except:
    st.markdown("<h2 style='text-align:center; color:#1E90FF;'>CLOUDTECH</h2>", unsafe_allow_html=True)

# === UPLOAD ===
file = st.file_uploader("Upload Excel", type="xlsx", key="file")

if file:
    with st.spinner("Processing..."):
        df = None
        for sheet in pd.ExcelFile(file).sheet_names:
            temp = pd.read_excel(file, sheet_name=sheet)
            cols = temp.columns.str.lower()
            if (cols.str.contains('number').any() and 
                cols.str.contains('created').any() and 
                cols.str.contains('actual.*end|work.*end', regex=True).any()):
                df = temp.copy()
                break

        if df is None:
            st.error("No valid data found.")
            st.stop()

        # Auto-detect
        c = df.columns.str.lower()
        df = df[[df.columns[c.str.contains('number')][0],
                 df.columns[c.str.contains('created')][0],
                 df.columns[c.str.contains('actual.*end|work.*end', regex=True)][0]]].copy()
        df.columns = ['Number', 'Created', 'End']

        # Dates
        def to_dt(x):
            try:
                if pd.notna(x) and x > 10000:
                    return pd.to_datetime(x, unit='D', origin='1899-12-30')
                return pd.to_datetime(x)
            except:
                return pd.NaT
        df['Created'] = df['Created'].apply(to_dt)
        df['End'] = df['End'].apply(to_dt)
        df = df.dropna()

        df['hours'] = (df['End'] - df['Created']).dt.total_seconds() / 3600
        df['STATUS'] = df['hours'] > 24

        within = len(df[~df['STATUS']])
        past = len(df[df['STATUS']])
        avg = df['hours'].mean()
        comp = within / len(df)

        # === METRICS ===
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(f"<div class='metric'>Within: <strong>{within}</strong></div>", unsafe_allow_html=True)
            st.markdown(f"<div class='metric'>Avg: <strong>{avg:.2f}h</strong></div>", unsafe_allow_html=True)
        with c2:
            st.markdown(f"<div class='metric'>Past: <strong>{past}</strong></div>", unsafe_allow_html=True)
            st.markdown(f"<div class='metric'>Compliance: <strong>{comp:.2%}</strong></div>", unsafe_allow_html=True)

        # === CHART ===
        fig, ax = plt.subplots()
        ax.pie([within, past], labels=['Within', 'Past'], autopct='%1.2f%%', colors=['#2ecc71', '#e74c3c'])
        buf = io.BytesIO()
        plt.savefig(buf, format='png', bbox_inches='tight', facecolor='#0a0e17')
        buf.seek(0)
        st.image(buf)

        # === EXCEL ===
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df[['Number', 'hours', 'STATUS']].to_excel(writer, 'Data', index=False)
            pd.DataFrame([
                ['Within SLA', within],
                ['Past SLA', past],
                ['Avg Hours', round(avg, 2)],
                ['Compliance', comp]
            ]).to_excel(writer, 'Data', index=False, header=False, startrow=len(df)+2)

            ws = writer.sheets['Data']
            for cell in ws[1]: cell.font = Font(bold=True)
            ws[f'B{len(df)+5}'].number_format = '0.00%'

        b64 = base64.b64encode(output.getvalue()).decode()
        st.markdown(f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="SLA_Report.xlsx" class="btn">Download Excel</a>', unsafe_allow_html=True)

else:
    st.info("Upload file to start")

st.markdown("<div class='footer'>© 2025 CloudTech</div>", unsafe_allow_html=True)