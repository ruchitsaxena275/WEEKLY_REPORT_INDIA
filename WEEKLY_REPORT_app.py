# streamlit_app.py
"""
Streamlit Storyboard app for JUNA PV weekly report.
Updated to avoid system-dependent libraries (uses matplotlib for PDF).
Requirements (example): streamlit, pandas, altair, matplotlib, python-docx, python-pptx, fpdf
"""

import streamlit as st
import pandas as pd
import altair as alt
import matplotlib.pyplot as plt
from io import BytesIO
from datetime import datetime, date
import re
from docx import Document
from pptx import Presentation
from pptx.util import Inches

st.set_page_config(page_title="JUNA PV — Weekly Storyboard", layout="wide", initial_sidebar_state="expanded")

# ---------- Helper functions ----------

def read_docx(path_or_file):
    """
    Accepts either a file-like object (from st.file_uploader) or a filesystem path.
    Returns a python-docx Document object.
    """
    if hasattr(path_or_file, "read"):
        return Document(path_or_file)
    else:
        return Document(path_or_file)

def docx_to_text(doc):
    """Return paragraphs list and list-of-tables (each table is list-of-rows, each row is list-of-cell-text)."""
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    tables = []
    for t in doc.tables:
        tbl = []
        for row in t.rows:
            tbl.append([cell.text.strip() for cell in row.cells])
        tables.append(tbl)
    return paragraphs, tables

def find_kpis(paragraphs):
    """Extract Weekly, MTD, YTD GWh heuristically from paragraphs."""
    kpis = {"weekly": None, "mtd": None, "ytd": None}
    text = "\n".join(paragraphs)
    # Try explicit pattern e.g., "Juna 9.93 25.19 206.41"
    m = re.search(r'Juna[\s:,-]*([\d]+\.[\d]+)[\s,;/-]*([\d]+\.[\d]+)[\s,;/-]*([\d]+\.[\d]+)', text, re.I)
    if m:
        kpis["weekly"], kpis["mtd"], kpis["ytd"] = m.group(1), m.group(2), m.group(3)
        return kpis
    # fallback: find three floats after a nearby 'Energy' token
    m2 = re.search(r'Energy[\s\S]{0,120}?([\d]+\.[\d]+)[\s\S]{0,20}?([\d]+\.[\d]+)[\s\S]{0,20}?([\d]+\.[\d]+)', text, re.I)
    if m2:
        kpis["weekly"], kpis["mtd"], kpis["ytd"] = m2.group(1), m2.group(2), m2.group(3)
    return kpis

def parse_breakdown_table(tables):
    """Return DataFrame if a breakdown-like table is found."""
    for tbl in tables:
        header = [c.lower() for c in tbl[0]]
        header_str = " ".join(header)
        if "start date" in header_str and ("breakdown" in header_str or "dc capacity" in header_str or "end date" in header_str):
            try:
                df = pd.DataFrame(tbl[1:], columns=tbl[0])
                return df
            except Exception:
                # fallback build with generic column names
                df = pd.DataFrame(tbl[1:])
                df.columns = [f"col_{i}" for i in range(len(df.columns))]
                return df
    return None

def parse_punch_points(tables):
    """Find punchpoint table like 'Block' 'Raised' 'Closed'."""
    for tbl in tables:
        header = [c.lower() for c in tbl[0]]
        header_str = " ".join(header)
        if "block" in header_str and ("raised" in header_str or "closed" in header_str or "balance" in header_str):
            try:
                df = pd.DataFrame(tbl[1:], columns=tbl[0])
                return df
            except Exception:
                df = pd.DataFrame(tbl[1:])
                df.columns = [f"col_{i}" for i in range(len(df.columns))]
                return df
    return None

def parse_robot_and_cleaning(paragraphs, tables):
    """Attempt to extract robot/module cleaning data from tables or paragraphs."""
    # prefer tables
    for tbl in tables:
        cols = [c.lower() for c in tbl[0]]
        if any("robot" in s or "module" in s or "clean" in s for s in cols):
            try:
                df = pd.DataFrame(tbl[1:], columns=tbl[0])
                return df, "table"
            except Exception:
                continue
    # fallback: scan paragraphs for 'Robot'/'Module Cleaning' block
    text = "\n".join(paragraphs)
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    start = None
    for i,l in enumerate(lines):
        if re.search(r'robot', l, re.I) or re.search(r'module cleaning', l, re.I) or re.search(r'robot trial', l, re.I):
            start = i
            break
    if start is not None:
        slice_lines = lines[start:start+10]
        return pd.DataFrame({"info": slice_lines}), "text"
    return None, None

def safe_float(s):
    try:
        return float(re.sub(r'[^\d.]','', str(s)))
    except:
        return None

def ensure_timeseries(df_ts):
    """Normalize timeseries DataFrame: ensure 'date' column and a numeric value column."""
    df = df_ts.copy()
    # Attempt to find a date-like column
    if 'date' not in [c.lower() for c in df.columns]:
        # try index 0
        df.columns = [str(c) for c in df.columns]
        if len(df.columns) >= 2:
            df = df.rename(columns={df.columns[0]:"date", df.columns[1]:"value"})
    else:
        # rename exact-case to 'date'
        cols = {c:c for c in df.columns}
        for c in df.columns:
            if c.lower() == 'date':
                cols[c] = 'date'
            elif c.lower() in ['active_power_mw','active power (mw)','activepower','value','active_power']:
                cols[c] = 'value'
        df = df.rename(columns=cols)
    # parse date
    try:
        df['date'] = pd.to_datetime(df['date'])
    except Exception:
        # try to coerce via pandas
        df['date'] = pd.to_datetime(df['date'], errors='coerce')
    # ensure a numeric column
    if 'value' not in df.columns and len(df.columns) >= 2:
        df = df.rename(columns={df.columns[1]:'value'})
    if 'value' in df.columns:
        df['value'] = pd.to_numeric(df['value'], errors='coerce')
    return df[['date','value']]

# ---------- PDF generator using matplotlib (pure python) ----------
def make_pdf_matplotlib(kpis, df_ts):
    """
    Create a simple one-page PDF using matplotlib. Returns BytesIO buffer.
    """
    buf = BytesIO()
    # A4 in inches approx
    fig = plt.figure(figsize=(8.27, 11.69))
    fig.suptitle("JUNA PV — Weekly Report", fontsize=16, fontweight='bold')

    # Header text with KPIs
    header_y = 0.92
    fig.text(0.05, header_y, f"Week: 20 Nov 2025 - 26 Nov 2025", fontsize=10)
    fig.text(0.05, header_y - 0.02, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", fontsize=8)
    fig.text(0.05, header_y - 0.06, f"Weekly Energy (GWh): {kpis.get('weekly','N/A')}", fontsize=11)
    fig.text(0.05, header_y - 0.09, f"MTD Energy (GWh): {kpis.get('mtd','N/A')}", fontsize=11)
    fig.text(0.05, header_y - 0.12, f"YTD Energy (GWh): {kpis.get('ytd','N/A')}", fontsize=11)

    # Plot timeseries in the lower half
    ax = fig.add_axes([0.08, 0.35, 0.88, 0.45])
    if df_ts is not None and not df_ts.empty:
        dfp = ensure_timeseries(df_ts)
        ax.plot(dfp['date'], dfp['value'], marker='o')
        ax.set_title("Weekly Active Power")
        ax.set_xlabel("Date")
        ax.set_ylabel("Active Power (MW)")
        fig.autofmt_xdate(rotation=30)
    else:
        ax.text(0.5, 0.5, "No timeseries data provided", ha='center', va='center')
    plt.tight_layout()
    fig.savefig(buf, format='pdf')
    plt.close(fig)
    buf.seek(0)
    return buf

# ---------- PPTX generator ----------
def make_pptx(kpis, df_ts, df_break=None, df_pp=None, df_robot=None):
    prs = Presentation()
    # Title
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "JUNA PV — Weekly Report"
    try:
        slide.placeholders[1].text = f"Week: 20 Nov 2025 - 26 Nov 2025\nGenerated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    except Exception:
        pass

    # KPI slide
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "KPIs"
    left = Inches(0.6)
    top = Inches(1.2)
    width = Inches(8)
    tx = s.shapes.add_textbox(left, top, width, Inches(1.6)).text_frame
    tx.clear()
    p = tx.add_paragraph()
    p.text = f"Weekly Energy (GWh): {kpis.get('weekly','N/A')}"
    p.level = 0
    p2 = tx.add_paragraph()
    p2.text = f"MTD Energy (GWh): {kpis.get('mtd','N/A')}"
    p2.level = 0
    p3 = tx.add_paragraph()
    p3.text = f"YTD Energy (GWh): {kpis.get('ytd','N/A')}"
    p3.level = 0

    # Timeseries as image
    img_stream = BytesIO()
    fig, ax = plt.subplots(figsize=(10,3))
    try:
        dfp = ensure_timeseries(df_ts)
        ax.plot(dfp['date'], dfp['value'])
    except Exception:
        ax.text(0.5, 0.5, "No timeseries", ha='center')
    ax.set_title("Weekly Active Power")
    ax.set_xlabel("Date")
    ax.set_ylabel("Active Power (MW)")
    fig.autofmt_xdate()
    fig.tight_layout()
    fig.savefig(img_stream, format='png')
    plt.close(fig)
    img_stream.seek(0)
    s.shapes.add_picture(img_stream, Inches(0.5), Inches(2.2), width=Inches(9))

    # Breakdown sample
    if df_break is not None:
        s2 = prs.slides.add_slide(prs.slide_layouts[5])
        s2.shapes.title.text = "Breakdown Log (sample)"
        txt = s2.shapes.add_textbox(Inches(0.4), Inches(1.2), Inches(9), Inches(4)).text_frame
        txt.clear()
        rows = df_break.head(6).to_dict(orient='records')
        for r in rows:
            line = " | ".join([f"{k}:{v}" for k, v in r.items()])
            p = txt.add_paragraph()
            p.text = line

    # Punchpoints sample
    if df_pp is not None:
        s3 = prs.slides.add_slide(prs.slide_layouts[5])
        s3.shapes.title.text = "Punch Points (sample)"
        txt = s3.shapes.add_textbox(Inches(0.4), Inches(1.2), Inches(9), Inches(4)).text_frame
        txt.clear()
        rows = df_pp.head(8).to_dict(orient='records')
        for r in rows:
            line = " | ".join([f"{k}:{v}" for k, v in r.items()])
            p = txt.add_paragraph()
            p.text = line

    # Robot sample
    if df_robot is not None:
        s4 = prs.slides.add_slide(prs.slide_layouts[5])
        s4.shapes.title.text = "Robot / Module Cleaning (sample)"
        txt = s4.shapes.add_textbox(Inches(0.4), Inches(1.2), Inches(9), Inches(4)).text_frame
        txt.clear()
        if isinstance(df_robot, pd.DataFrame):
            rows = df_robot.head(8).to_dict(orient='records')
            for r in rows:
                line = " | ".join([f"{k}:{v}" for k, v in r.items()])
                p = txt.add_paragraph()
                p.text = line
        else:
            p = txt.add_paragraph()
            p.text = str(df_robot)

    out = BytesIO()
    prs.save(out)
    out.seek(0)
    return out

# ---------- UI ----------

st.title("JUNA PV — Weekly Storyboard")
st.markdown("Create an attractive weekly report from your DOCX. (This instance uses an embedded file if present.)")

# Sidebar: upload or use embedded file
st.sidebar.header("Data source")
uploaded = st.sidebar.file_uploader("Upload DOCX (optional) — will override embedded file", type=["docx"])
use_embedded_btn = st.sidebar.button("Use embedded uploaded weekly report (if present on server)")

default_path = "/mnt/data/2025 India -Juna Weekly report_Week 47.docx"
doc_source = None
if uploaded:
    doc_source = uploaded
elif use_embedded_btn:
    # try to use the file present at default_path
    try:
        open(default_path, "rb").close()
        doc_source = default_path
    except Exception:
        st.sidebar.warning("Embedded file not found at /mnt/data. Please upload a DOCX.")
        doc_source = None
else:
    # if embedded exists, use it by default (convenience)
    try:
        open(default_path, "rb").close()
        doc_source = default_path
    except Exception:
        doc_source = None

if doc_source is None:
    st.info("Upload the weekly DOCX report (or click 'Use embedded' if you placed it in /mnt/data).")
    st.stop()

# Read docx
try:
    doc = read_docx(doc_source)
    paragraphs, tables = docx_to_text(doc)
except Exception as e:
    st.error(f"Failed to read DOCX: {e}")
    st.stop()

# Parse KPIs
kpis = find_kpis(paragraphs)
weekly_energy = kpis.get("weekly") or "N/A"
mtd_energy = kpis.get("mtd") or "N/A"
ytd_energy = kpis.get("ytd") or "N/A"

# KPI band
col1, col2, col3, col4 = st.columns(4)
col1.metric("Weekly Energy (GWh)", weekly_energy)
col2.metric("MTD Energy (GWh)", mtd_energy)
col3.metric("YTD Energy (GWh)", ytd_energy)
plant_status = "Operational (curtailed)" if any("curtail" in p.lower() or "nrl" in p.lower() for p in paragraphs) else "Operational"
col4.metric("Plant status", plant_status)

st.markdown("---")

# Timeseries input (optional)
st.subheader("Weekly Active Power")
timeseries_csv = st.file_uploader("Upload daily timeseries CSV (columns: date, value OR date, active_power_mw) (optional)", type=["csv"])
if timeseries_csv:
    try:
        df_ts = pd.read_csv(timeseries_csv)
    except Exception as e:
        st.error("Failed to read CSV: " + str(e))
        df_ts = pd.DataFrame({"date": pd.date_range(date(2025,11,20), date(2025,11,26)), "value":[140,138,130,135,132,120,126]})
else:
    # placeholder simple time series
    df_ts = pd.DataFrame({"date": pd.date_range("2025-11-20", "2025-11-26"), "value":[140,138,130,135,132,120,126]})

# Show chart
try:
    df_plot = ensure_timeseries(df_ts)
    chart = alt.Chart(df_plot).mark_area(opacity=0.45).encode(
        x=alt.X("date:T", title="Date"),
        y=alt.Y("value:Q", title="Active Power (MW)"),
        tooltip=["date:T", "value:Q"]
    ).interactive()
    st.altair_chart(chart, use_container_width=True)
except Exception as e:
    st.write("Unable to render chart:", e)

st.markdown("---")

# Two-column content
left, right = st.columns([1,2])
with left:
    st.subheader("Snapshot / Highlights")
    highlights = []
    for p in paragraphs:
        low = p.lower()
        if any(kw in low for kw in ["thermography", "robot", "curtail", "system integration", "module cleaning", "cable theft", "hall ct"]):
            highlights.append(p)
    if not highlights:
        for i,p in enumerate(paragraphs):
            if p.strip().lower().startswith("highlights"):
                highlights = paragraphs[i:i+8]
                break
    if highlights:
        for h in highlights:
            st.write("- " + h)
    else:
        st.write("No clear highlights parsed.")
    st.markdown("**Quick filters**")
    st.selectbox("Block", ["All"] + [f"Block-{i}" for i in range(1,21)])
    st.date_input("Start date", value=date(2025,11,20))
    st.date_input("End date", value=date(2025,11,26))

with right:
    st.subheader("Timeline — This Week")
    timeline = []
    for p in paragraphs:
        if re.search(r'\b(20th|21st|22nd|23rd|24th|25th|26th)\b', p, re.I) or re.search(r'\b20-Nov|\b21-Nov|\b26-Nov', p, re.I):
            timeline.append(p)
    if not timeline:
        timeline = [p for p in paragraphs if "half yearly" in p.lower() or "cable theft" in p.lower() or "hall ct" in p.lower()]
    if timeline:
        for ev in timeline:
            st.info(ev)
    else:
        st.write("No timeline events parsed.")

st.markdown("---")

# Breakdown table
st.subheader("Equipment Breakdown Log")
df_break = parse_breakdown_table(tables)
if df_break is not None:
    st.dataframe(df_break)
    st.download_button("Download breakdown CSV", df_break.to_csv(index=False).encode('utf-8'), "breakdown_log.csv", "text/csv")
else:
    st.write("No breakdown table auto-detected. You can upload a CSV/Excel for breakdowns.")
    uploaded_break = st.file_uploader("Upload breakdown CSV/XLSX", type=["csv","xlsx"], key="breakupload")
    if uploaded_break:
        try:
            if uploaded_break.name.lower().endswith('.csv'):
                df_break = pd.read_csv(uploaded_break)
            else:
                df_break = pd.read_excel(uploaded_break)
            st.dataframe(df_break)
            st.download_button("Download breakdown CSV (parsed)", df_break.to_csv(index=False).encode('utf-8'), "breakdown_log.csv", "text/csv")
        except Exception as e:
            st.error("Failed to read breakdown file: " + str(e))

# Punch points
st.subheader("Punch Points Summary")
df_pp = parse_punch_points(tables)
if df_pp is not None:
    st.dataframe(df_pp)
    # try to chart if columns present
    try:
        first_col = df_pp.columns[0]
        numeric_cols = [c for c in df_pp.columns if any(k in c.lower() for k in ['raised','closed','balance'])]
        if numeric_cols:
            chart_df = df_pp[[first_col] + numeric_cols].copy()
            for c in numeric_cols:
                chart_df[c] = pd.to_numeric(chart_df[c].astype(str).str.replace(r'[^\d.]','', regex=True), errors='coerce')
            chart_df = chart_df.set_index(first_col)
            st.bar_chart(chart_df)
    except Exception:
        pass
else:
    st.write("No punchpoints table auto-detected.")

# Robot/module cleaning logs
st.markdown("---")
st.subheader("Robot & Module Cleaning Logs")
df_robot, rtype = parse_robot_and_cleaning(paragraphs, tables)
if df_robot is not None:
    st.dataframe(df_robot)
    if isinstance(df_robot, pd.DataFrame):
        st.download_button("Download Robot/Module CSV", df_robot.to_csv(index=False).encode('utf-8'), "robot_module.csv", "text/csv")
else:
    st.write("No robot/module cleaning entries auto-detected.")

st.markdown("---")
st.subheader("Export / Downloads")

# CSV downloads
if 'df_break' in locals() and df_break is not None:
    st.download_button("Download breakdown CSV", df_break.to_csv(index=False).encode('utf-8'), "breakdown_log.csv", "text/csv")
if df_pp is not None:
    st.download_button("Download punchpoints CSV", df_pp.to_csv(index=False).encode('utf-8'), "punchpoints.csv", "text/csv")

# PPTX export
if st.button("Generate PPTX"):
    try:
        pptx_buf = make_pptx(kpis, df_ts, df_break, df_pp, df_robot)
        st.success("PPTX generated.")
        st.download_button("Download PPTX", data=pptx_buf.getvalue(), file_name="juna_weekly_report.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    except Exception as e:
        st.error("Failed to generate PPTX: " + str(e))

# PDF export (matplotlib-backed)
if st.button("Generate PDF (simple)"):
    try:
        pdf_buf = make_pdf_matplotlib(kpis, df_ts)
        st.success("PDF generated.")
        st.download_button("Download PDF", data=pdf_buf.getvalue(), file_name="juna_weekly_report.pdf", mime="application/pdf")
    except Exception as e:
        st.error("Failed to generate PDF: " + str(e))

st.markdown("---")
st.caption("If a section is not parsed correctly, upload the relevant CSV/Excel (timeseries, breakdown, punchpoints).")
