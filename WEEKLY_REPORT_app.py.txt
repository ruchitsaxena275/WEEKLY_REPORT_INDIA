# streamlit_app.py
import streamlit as st
import pandas as pd
import altair as alt
import matplotlib.pyplot as plt
from io import BytesIO
from datetime import datetime
import re
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

st.set_page_config(page_title="JUNA PV — Weekly Storyboard", layout="wide")

# ---------- Helpers ----------
def read_docx(path_or_file):
    """
    Accepts either a file-like object (from st.file_uploader) or a filesystem path.
    Returns Document object.
    """
    if hasattr(path_or_file, "read"):
        # file-like from uploader
        return Document(path_or_file)
    else:
        return Document(path_or_file)

def docx_to_text(doc):
    """Return full text and tables list for simpler parsing."""
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    tables = []
    for t in doc.tables:
        # convert table to list of lists
        tbl = []
        for row in t.rows:
            tbl.append([cell.text.strip() for cell in row.cells])
        tables.append(tbl)
    return paragraphs, tables

def find_kpis(paragraphs):
    """Extract Energy KPIs using regex heuristics."""
    kpis = {"weekly": None, "mtd": None, "ytd": None}
    text = "\n".join(paragraphs)
    # find patterns like 'PV\nJuna\n9.93\n25.19\n206.41' or labelled 'Energy   GWh'
    # Search for three floats that appear near 'Energy' and 'Juna'
    m = re.search(r'Juna\s*([\d\.]+)\s*([\d\.]+)\s*([\d\.]+)', text, re.I)
    if m:
        kpis["weekly"], kpis["mtd"], kpis["ytd"] = m.group(1), m.group(2), m.group(3)
        return kpis
    # fallback: find three nearest floats after 'Energy'
    m2 = re.search(r'Energy\s*GWh.*?([\d\.]+).*?([\d\.]+).*?([\d\.]+)', text, re.S | re.I)
    if m2:
        kpis["weekly"], kpis["mtd"], kpis["ytd"] = m2.group(1), m2.group(2), m2.group(3)
    return kpis

def parse_breakdown_table(tables):
    """
    Look for a table whose header contains 'Start date' and 'Breakdown time' etc.
    Return a pandas DataFrame or None.
    """
    for tbl in tables:
        header = [c.lower() for c in tbl[0]]
        header_str = " ".join(header)
        if "start date" in header_str and ("breakdown" in header_str or "dc capacity" in header_str):
            df = pd.DataFrame(tbl[1:], columns=tbl[0])
            return df
    return None

def parse_punch_points(tables):
    """Find the table that contains 'BLOCK NO.' or 'Block-1' style rows."""
    for tbl in tables:
        header = [c.lower() for c in tbl[0]]
        header_str = " ".join(header)
        if "block" in header_str and ("raised" in header_str or "closed" in header_str):
            df = pd.DataFrame(tbl[1:], columns=tbl[0])
            return df
    # fallback: detect long table with Block- entries inside rows
    for tbl in tables:
        rows_text = " ".join([" ".join(r).lower() for r in tbl])
        if "block-1" in rows_text or "block-2" in rows_text:
            df = pd.DataFrame(tbl[1:], columns=tbl[0])
            return df
    return None

def parse_robot_and_cleaning(paragraphs, tables):
    """Extract robot/cleaning logs from paragraphs or tables."""
    text = "\n".join(paragraphs)
    # Try to extract the 'Robot trial' and 'Module Cleaning Work' sections which are often in a table
    # Find table with two columns 'Date' and 'Activity' or similar
    for tbl in tables:
        cols = [c.lower() for c in tbl[0]]
        if ("date" in cols[0].lower() and ("activity" in cols[1].lower() or "block" in cols[1].lower())) or ("activity" in " ".join(cols).lower()):
            df = pd.DataFrame(tbl[1:], columns=tbl[0])
            # try to detect which type: Module Cleaning or Robot Trail
            if any("Module" in str(x) or "Cleaning" in str(x) for x in df.values.flatten()):
                return df, "module_cleaning"
            if any("Robot" in str(x) for x in df.values.flatten()):
                return df, "robot_trail"
    # else search paragraphs for 'Robot Trail' headings and parse following lines
    lines = [ln for ln in text.splitlines() if ln.strip()]
    robot_lines = []
    recording = False
    for ln in lines:
        if 'robot' in ln.lower() or 'module cleaning' in ln.lower():
            recording = True
        if recording:
            robot_lines.append(ln)
            # heuristics: stop after an empty line or 'Highlights' found
            if ln.strip().lower().startswith('highlights'):
                break
    # return minimal structure if found
    if robot_lines:
        return pd.DataFrame({"raw": robot_lines}), "text_table"
    return None, None

def safe_float(s):
    try:
        return float(re.sub(r'[^\d\.]', '', str(s)))
    except:
        return None

# ---------- UI ----------
st.title("JUNA PV — Weekly Storyboard (Week report parser + exports)")

st.markdown("**Source document:** 2025 India - Juna Weekly report (Week 47). :contentReference[oaicite:1]{index=1}")

# Allow user to upload alternative docx or use embedded one
st.sidebar.header("Data source")
use_uploaded = st.sidebar.file_uploader("Upload DOCX (optional) — will override default", type=["docx"])
use_default_btn = st.sidebar.button("Use embedded uploaded weekly report (the file you provided)")

# Determine doc source path - fallback to embedded file path in /mnt/data when available
default_path = "/mnt/data/2025 India -Juna Weekly report_Week 47.docx"
doc_source = None
if use_uploaded:
    doc_source = use_uploaded
elif use_default_btn:
    doc_source = default_path
else:
    # by default load embedded file if exists
    try:
        # try to open embedded path
        open(default_path, "rb").close()
        doc_source = default_path
    except Exception:
        st.sidebar.warning("No default file found. Upload a DOCX to proceed.")
        doc_source = None

if not doc_source:
    st.info("Upload the weekly DOCX report (or press 'Use embedded' in sidebar if you uploaded the file earlier).")
    st.stop()

# Read the document
try:
    doc = read_docx(doc_source)
    paragraphs, tables = docx_to_text(doc)
except Exception as e:
    st.error(f"Unable to read DOCX: {e}")
    st.stop()

# Parse KPIs
kpis = find_kpis(paragraphs)
weekly_energy = kpis.get("weekly") or "N/A"
mtd_energy = kpis.get("mtd") or "N/A"
ytd_energy = kpis.get("ytd") or "N/A"

# Top KPI band
kcol1, kcol2, kcol3, kcol4 = st.columns(4)
kcol1.metric("Weekly Energy (GWh)", weekly_energy)
kcol2.metric("MTD Energy (GWh)", mtd_energy)
kcol3.metric("YTD Energy (GWh)", ytd_energy)
# Derive plant status from paragraphs heuristics
plant_status = "Operational (curtailed as per NRLDC)" if any("curtail" in p.lower() or "nrl" in p.lower() for p in paragraphs) else "Operational"
kcol4.metric("Plant status", plant_status)

st.markdown("---")

# Hero: Weekly active power chart area
st.subheader("Weekly Active Power")
# Try to extract a line with dates or a small ascii chart — fallback: placeholder series
# If the doc doesn't contain daily numbers we allow CSV upload for precise timeseries
timeseries_csv = st.file_uploader("Upload daily timeseries CSV (columns: date, active_power_mw) to populate the chart (optional)", type=["csv"])
if timeseries_csv:
    df_ts = pd.read_csv(timeseries_csv, parse_dates=["date"])
else:
    # attempt to infer from paragraph lines; simple heuristic: look for lines containing 20-Nov ... 26-nov
    # fallback to placeholder synthetic data to visualize layout
    df_ts = pd.DataFrame({
        "date": pd.date_range("2025-11-20", "2025-11-26"),
        "active_power_mw": [140, 138, 130, 135, 132, 120, 126]
    })
# interactive Altair chart
chart = alt.Chart(df_ts).mark_area(opacity=0.45).encode(
    x=alt.X("date:T", title="Date"),
    y=alt.Y("active_power_mw:Q", title="Active Power (MW)"),
    tooltip=["date:T", "active_power_mw:Q"]
).interactive()
st.altair_chart(chart, use_container_width=True)

st.markdown("---")

# Two-column main area
left, right = st.columns([1,2])

with left:
    st.subheader("Snapshot / Highlights")
    # Show highlights found
    highlights = []
    for p in paragraphs:
        low = p.lower()
        if any(keyword in low for keyword in ["thermography", "robot", "curtail", "no other major", "system integration", "node tracker", "220kv"]):
            highlights.append(p)
    if not highlights:
        # try to find block of text labeled Highlights
        for i, p in enumerate(paragraphs):
            if p.strip().lower().startswith("highlights"):
                # include subsequent lines
                highlights = paragraphs[i:i+6]
                break
    if highlights:
        for h in highlights:
            st.write("- " + h)
    else:
        st.write("No automated highlights found in the document.")
    st.markdown("**Quick filters**")
    st.selectbox("Block filter", ["All"] + [f"Block-{i}" for i in range(1,21)])
    st.date_input("Start date", value=datetime(2025,11,20))
    st.date_input("End date", value=datetime(2025,11,26))
    st.markdown("**Site notes**")
    # show a few first paragraphs as site notes
    for p in paragraphs[:6]:
        st.write(p)

with right:
    st.subheader("Timeline — This Week")
    # find lines mentioning dates and events
    timeline = []
    for p in paragraphs:
        if re.search(r'\b(20th|21st|22nd|23rd|24th|25th|26th)\b', p, re.I) or re.search(r'\b20-Nov|\b21-Nov|\b26-Nov', p, re.I):
            timeline.append(p)
    if not timeline:
        # search for "half yearly Maintenance" lines
        timeline = [p for p in paragraphs if "half yearly" in p.lower() or "half-yearly" in p.lower() or "cable theft" in p.lower()]
    if timeline:
        for ev in timeline:
            st.info(ev)
    else:
        st.write("No timeline events parsed automatically.")

st.markdown("---")

# Breakdown table parsing
st.subheader("Equipment Breakdown Log")
df_break = parse_breakdown_table(tables)
if df_break is not None:
    # display and allow filtering
    st.dataframe(df_break)
    st.download_button("Download breakdown CSV", df_break.to_csv(index=False).encode('utf-8'), "breakdown_log.csv", "text/csv")
else:
    st.write("No structured breakdown table found automatically. If you have a CSV/Excel of breakdowns, upload it below.")
    uploaded_break_csv = st.file_uploader("Upload breakdown CSV", type=["csv","xlsx"])
    if uploaded_break_csv:
        try:
            if uploaded_break_csv.name.lower().endswith(".csv"):
                df_break = pd.read_csv(uploaded_break_csv)
            else:
                df_break = pd.read_excel(uploaded_break_csv)
            st.dataframe(df_break)
            st.download_button("Download breakdown CSV (parsed)", df_break.to_csv(index=False).encode('utf-8'), "breakdown_log.csv", "text/csv")
        except Exception as e:
            st.error(f"Error reading uploaded breakdown file: {e}")

# Punch points
st.subheader("ITC Area Punch Points")
df_pp = parse_punch_points(tables)
if df_pp is not None:
    # coerce numeric
    try:
        df_pp2 = df_pp.copy()
        # normalize header names
        df_pp2.columns = [c.strip() for c in df_pp2.columns]
        if 'BLOCK NO.' in df_pp2.columns or any('Block' in c for c in df_pp2.columns):
            # attempt to reshape to columns block, raised, closed
            # If table is long-format (Block-1 value rows), try to parse accordingly
            # If already in columns, just show
            st.dataframe(df_pp2)
            # chart
            show_cols = [c for c in df_pp2.columns if any(x in c.lower() for x in ['raised','closed','balance'])]
            if show_cols:
                chart_df = df_pp2[[df_pp2.columns[0]] + show_cols].copy()
                # convert numeric
                for c in show_cols:
                    chart_df[c] = chart_df[c].apply(safe_float)
                chart_df = chart_df.set_index(df_pp2.columns[0])
                st.bar_chart(chart_df)
        else:
            st.dataframe(df_pp2)
    except Exception as e:
        st.write("Punch points parsed, but unable to render chart:", e)
else:
    st.write("No punch-points table detected; total punch points in doc (if present) may be parsed into text.")

# Robot & Module Cleaning
st.markdown("---")
st.subheader("Robot & Module Cleaning Logs")
df_robot, robot_type = parse_robot_and_cleaning(paragraphs, tables)
if df_robot is not None:
    st.dataframe(df_robot)
else:
    st.write("No robot/module cleaning table found automatically.")

# Downloads / Exports
st.markdown("---")
st.subheader("Export the Storyboard / Raw Data")

# CSV exports for parsed tables
if df_break is not None:
    st.download_button("Download breakdown CSV", df_break.to_csv(index=False).encode('utf-8'), "breakdown_log.csv", "text/csv")
if df_pp is not None:
    st.download_button("Download punchpoints CSV", df_pp.to_csv(index=False).encode('utf-8'), "punchpoints.csv", "text/csv")
if df_robot is not None:
    # df_robot might be DataFrame
    if isinstance(df_robot, pd.DataFrame):
        st.download_button("Download robot/module CSV", df_robot.to_csv(index=False).encode('utf-8'), "robot_module.csv", "text/csv")

# PPTX generation
def make_pptx(kpis, df_ts, df_break=None, df_pp=None, df_robot=None):
    prs = Presentation()
    # Title slide
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = "JUNA PV — Weekly Report"
    subtitle.text = f"Week: 20 Nov 2025 - 26 Nov 2025\nGenerated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
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
    # Add chart image for timeseries
    img_stream = BytesIO()
    fig, ax = plt.subplots(figsize=(10,3))
    ax.plot(df_ts['date'], df_ts[df_ts.columns[1]])
    ax.set_title("Weekly Active Power")
    ax.set_xlabel("Date")
    ax.set_ylabel("Active Power (MW)")
    fig.autofmt_xdate()
    plt.tight_layout()
    fig.savefig(img_stream, format='png')
    plt.close(fig)
    img_stream.seek(0)
    pic = s.shapes.add_picture(img_stream, Inches(0.5), Inches(2.2), width=Inches(9))
    # Breakdown slide
    if df_break is not None:
        s2 = prs.slides.add_slide(prs.slide_layouts[5])
        s2.shapes.title.text = "Breakdown Log (sample)"
        # add first 6 rows as text
        tbl_txt = s2.shapes.add_textbox(Inches(0.4), Inches(1.2), Inches(9), Inches(4)).text_frame
        tbl_txt.clear()
        rows_to_show = df_break.head(6).to_dict(orient='records')
        for r in rows_to_show:
            line = " | ".join([f"{k}:{v}" for k, v in r.items()])
            p = tbl_txt.add_paragraph()
            p.text = line
    # Punch points slide
    if df_pp is not None:
        s3 = prs.slides.add_slide(prs.slide_layouts[5])
        s3.shapes.title.text = "Punch Points (sample)"
        tbl_txt = s3.shapes.add_textbox(Inches(0.4), Inches(1.2), Inches(9), Inches(4)).text_frame
        tbl_txt.clear()
        rows_to_show = df_pp.head(10).to_dict(orient='records')
        for r in rows_to_show:
            line = " | ".join([f"{k}:{v}" for k, v in r.items()])
            p = tbl_txt.add_paragraph()
            p.text = line
    # Robot slide
    if df_robot is not None:
        s4 = prs.slides.add_slide(prs.slide_layouts[5])
        s4.shapes.title.text = "Robot / Cleaning (sample)"
        tbl_txt = s4.shapes.add_textbox(Inches(0.4), Inches(1.2), Inches(9), Inches(4)).text_frame
        tbl_txt.clear()
        if isinstance(df_robot, pd.DataFrame):
            rows_to_show = df_robot.head(10).to_dict(orient='records')
            for r in rows_to_show:
                line = " | ".join([f"{k}:{v}" for k, v in r.items()])
                p = tbl_txt.add_paragraph()
                p.text = line
        else:
            p = tbl_txt.add_paragraph()
            p.text = str(df_robot)
    # Return bytes
    out = BytesIO()
    prs.save(out)
    out.seek(0)
    return out

if st.button("Generate PPTX report (Download)"):
    try:
        pptx_bytes = make_pptx(kpis, df_ts, df_break, df_pp, df_robot)
        st.success("PPTX generated. Click download below.")
        st.download_button("Download PPTX", data=pptx_bytes.getvalue(), file_name="juna_weekly_report.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    except Exception as e:
        st.error(f"Failed to generate PPTX: {e}")

# Simple PDF generation (text summary + chart)
def make_pdf(kpis, df_ts):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    width, height = A4
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, height - 50, "JUNA PV — Weekly Report")
    c.setFont("Helvetica", 11)
    c.drawString(50, height - 70, f"Week: 20 Nov 2025 - 26 Nov 2025")
    c.drawString(50, height - 90, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    # KPIs
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, height - 120, "KPIs:")
    c.setFont("Helvetica", 11)
    c.drawString(60, height - 140, f"Weekly Energy (GWh): {kpis.get('weekly','N/A')}")
    c.drawString(60, height - 155, f"MTD Energy (GWh): {kpis.get('mtd','N/A')}")
    c.drawString(60, height - 170, f"YTD Energy (GWh): {kpis.get('ytd','N/A')}")
    # Add chart image
    img_stream = BytesIO()
    fig, ax = plt.subplots(figsize=(6,2.5))
    ax.plot(df_ts['date'], df_ts[df_ts.columns[1]])
    ax.set_title("Weekly Active Power")
    ax.set_xlabel("Date")
    ax.set_ylabel("Active Power (MW)")
    fig.autofmt_xdate()
    plt.tight_layout()
    fig.savefig(img_stream, format='png', dpi=150)
    plt.close(fig)
    img_stream.seek(0)
    # drawImage requires a filename-like object
    c.drawImage(img_stream, 50, height - 360, width=500, height=140)
    c.showPage()
    c.save()
    buf.seek(0)
    return buf

if st.button("Generate PDF (simple)"):
    try:
        pdf_buf = make_pdf(kpis, df_ts)
        st.success("PDF generated.")
        st.download_button("Download PDF", data=pdf_buf.getvalue(), file_name="juna_weekly_report.pdf", mime="application/pdf")
    except Exception as e:
        st.error(f"Failed to generate PDF: {e}")

st.markdown("---")
st.caption("This app parsed the DOCX you uploaded. If a section was not auto-detected, upload the specific CSV/Excel for that section (timeseries, breakdown, punchpoints).")
