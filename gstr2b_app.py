import streamlit as st
import openpyxl
from openpyxl import Workbook
from copy import copy
import os
import re
import io
import zipfile
from collections import defaultdict

# ─────────────────────────────────────────────
#  PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="GSTR-2B Compiler",
    page_icon="📊",
    layout="centered",
)

# ─────────────────────────────────────────────
#  CUSTOM CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');

html, body, [class*="css"] {
    font-family: 'IBM Plex Sans', sans-serif;
}

/* Dark background overall */
.stApp {
    background-color: #0f1117;
    color: #e8eaf0;
}

/* Header banner */
.header-banner {
    background: linear-gradient(135deg, #1a3a5c 0%, #0d2137 60%, #071525 100%);
    border: 1px solid #1e4976;
    border-radius: 12px;
    padding: 2rem 2.5rem;
    margin-bottom: 2rem;
    position: relative;
    overflow: hidden;
}
.header-banner::before {
    content: '';
    position: absolute;
    top: -40px; right: -40px;
    width: 160px; height: 160px;
    border-radius: 50%;
    background: rgba(30,120,200,0.08);
}
.header-banner h1 {
    font-size: 1.9rem;
    font-weight: 700;
    color: #ffffff;
    margin: 0 0 0.3rem 0;
    letter-spacing: -0.5px;
}
.header-banner p {
    font-size: 0.92rem;
    color: #8aadcc;
    margin: 0;
    font-weight: 300;
}
.header-badge {
    display: inline-block;
    background: rgba(30,180,100,0.15);
    border: 1px solid rgba(30,180,100,0.4);
    color: #4dca8a;
    font-size: 0.7rem;
    font-weight: 600;
    letter-spacing: 1px;
    text-transform: uppercase;
    padding: 2px 10px;
    border-radius: 20px;
    margin-bottom: 0.8rem;
    font-family: 'IBM Plex Mono', monospace;
}

/* Info cards */
.info-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 0.8rem;
    margin-bottom: 1.5rem;
}
.info-card {
    background: #161b27;
    border: 1px solid #1e2d42;
    border-radius: 8px;
    padding: 1rem;
    text-align: center;
}
.info-card .num {
    font-size: 1.6rem;
    font-weight: 700;
    color: #3b9eff;
    font-family: 'IBM Plex Mono', monospace;
    line-height: 1;
}
.info-card .label {
    font-size: 0.72rem;
    color: #6b7fa0;
    margin-top: 0.3rem;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

/* Section labels */
.section-label {
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    color: #3b9eff;
    margin-bottom: 0.5rem;
    font-family: 'IBM Plex Mono', monospace;
}

/* File list */
.file-list {
    background: #161b27;
    border: 1px solid #1e2d42;
    border-radius: 8px;
    padding: 0.8rem 1rem;
    margin-bottom: 1rem;
    max-height: 200px;
    overflow-y: auto;
}
.file-item {
    display: flex;
    align-items: center;
    gap: 0.5rem;
    padding: 0.25rem 0;
    font-size: 0.82rem;
    color: #b0c4de;
    font-family: 'IBM Plex Mono', monospace;
    border-bottom: 1px solid #1a2235;
}
.file-item:last-child { border-bottom: none; }
.file-dot {
    width: 6px; height: 6px;
    border-radius: 50%;
    background: #3b9eff;
    flex-shrink: 0;
}

/* Log box */
.log-box {
    background: #0a0e17;
    border: 1px solid #1e2d42;
    border-radius: 8px;
    padding: 1rem;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.78rem;
    color: #4dca8a;
    max-height: 260px;
    overflow-y: auto;
    line-height: 1.7;
}
.log-box .log-dim { color: #3a5070; }
.log-box .log-warn { color: #f0a500; }
.log-box .log-head { color: #3b9eff; }

/* Month tag */
.month-tag {
    display: inline-block;
    background: #0d2137;
    border: 1px solid #1e4976;
    color: #5aadff;
    padding: 2px 8px;
    border-radius: 4px;
    font-family: 'IBM Plex Mono', monospace;
    font-size: 0.78rem;
    margin: 2px;
}

/* Summary box */
.summary-box {
    background: linear-gradient(135deg, #0d2a1a, #071a10);
    border: 1px solid #1a5c30;
    border-radius: 10px;
    padding: 1.2rem 1.5rem;
    margin: 1.2rem 0;
}
.summary-box h4 {
    color: #4dca8a;
    margin: 0 0 0.6rem 0;
    font-size: 0.9rem;
    font-weight: 600;
}
.summary-row {
    display: flex;
    justify-content: space-between;
    font-size: 0.82rem;
    padding: 0.2rem 0;
    color: #a0c8b0;
    border-bottom: 1px solid #0f3d20;
}
.summary-row:last-child { border-bottom: none; }
.summary-val {
    font-family: 'IBM Plex Mono', monospace;
    color: #ffffff;
}

/* Streamlit overrides */
.stFileUploader > div {
    background: #161b27 !important;
    border: 1.5px dashed #1e4976 !important;
    border-radius: 10px !important;
}
.stButton > button {
    background: linear-gradient(135deg, #1a6fc4, #0d4f8c) !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 0.6rem 2rem !important;
    font-family: 'IBM Plex Sans', sans-serif !important;
    font-weight: 600 !important;
    font-size: 0.95rem !important;
    letter-spacing: 0.3px !important;
    width: 100% !important;
    transition: opacity 0.2s !important;
}
.stButton > button:hover { opacity: 0.88 !important; }
.stDownloadButton > button {
    background: linear-gradient(135deg, #1a7a45, #0f5530) !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 0.6rem 2rem !important;
    font-family: 'IBM Plex Sans', sans-serif !important;
    font-weight: 600 !important;
    font-size: 0.95rem !important;
    width: 100% !important;
}
div[data-testid="stProgress"] > div > div {
    background: linear-gradient(90deg, #1a6fc4, #4dca8a) !important;
}
.stSuccess { background: #0d2a1a !important; border-color: #1a5c30 !important; }
.stWarning { background: #2a1e0a !important; border-color: #5c3d0a !important; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  CONSTANTS
# ─────────────────────────────────────────────
SUMMARY_SHEETS = {'Read me', 'ITC Available', 'ITC not available', 'ITC Rejected'}

METADATA_ROWS = {
    'B2B': 6,               'B2BA': 7,
    'B2B-CDNR': 6,          'B2B-CDNRA': 7,
    'ECO': 6,               'ECOA': 7,
    'ISD': 6,               'ISDA': 7,
    'IMPG': 6,              'IMPGSEZ': 6,
    'B2B(Rejected)': 6,     'B2BA(Rejected)': 7,
    'B2B-CDNR(Rejected)': 6,'B2B-CDNRA(Rejected)': 7,
    'ECO(Rejected)': 6,     'ECOA(Rejected)': 7,
}

MONTH_NAMES = {
    '01':'January','02':'February','03':'March','04':'April',
    '05':'May','06':'June','07':'July','08':'August',
    '09':'September','10':'October','11':'November','12':'December'
}

# ─────────────────────────────────────────────
#  CORE FUNCTIONS
# ─────────────────────────────────────────────
def get_metadata_rows(sheet_name):
    return METADATA_ROWS.get(sheet_name, 6)

def copy_cell_style(src, dst):
    if src.has_style:
        dst.font        = copy(src.font)
        dst.fill        = copy(src.fill)
        dst.border      = copy(src.border)
        dst.alignment   = copy(src.alignment)
        dst.number_format = src.number_format
        dst.protection  = copy(src.protection)

def copy_metadata(src_ws, dst_ws, num_rows):
    for row in src_ws.iter_rows(min_row=1, max_row=num_rows):
        for sc in row:
            dc = dst_ws.cell(row=sc.row, column=sc.column)
            dc.value = sc.value
            copy_cell_style(sc, dc)
    for mr in src_ws.merged_cells.ranges:
        if mr.max_row <= num_rows:
            dst_ws.merge_cells(str(mr))
    for r in range(1, num_rows + 1):
        if r in src_ws.row_dimensions:
            dst_ws.row_dimensions[r].height = src_ws.row_dimensions[r].height
    for col, cd in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[col].width = cd.width

def append_data(src_ws, dst_ws, skip_rows):
    if src_ws.max_row <= skip_rows:
        return 0
    count = 0
    for row in src_ws.iter_rows(min_row=skip_rows + 1, values_only=True):
        dst_ws.append(list(row))
        count += 1
    return count

def get_fy_sort_key(mp):
    mm, yyyy = int(mp[:2]), int(mp[2:])
    fiscal_month = (mm - 3) if mm >= 4 else (mm + 9)
    fiscal_year  = yyyy if mm >= 4 else yyyy - 1
    return (fiscal_year, fiscal_month)

def month_label(mp):
    mm, yyyy = mp[:2], mp[2:]
    return f"{MONTH_NAMES.get(mm, mm)}-{yyyy}"

def group_files(uploaded_files):
    pattern = re.compile(r'^(\d{6})_.*\.xlsx$', re.IGNORECASE)
    groups  = defaultdict(list)
    for uf in uploaded_files:
        if pattern.match(uf.name):
            mp = uf.name[:6]
            groups[mp].append(uf)
        else:
            st.warning(f"⚠️ Skipped (unrecognised name): {uf.name}")

    def file_sort_key(uf):
        num = re.search(r'_(\d+)\.xlsx$', uf.name, re.IGNORECASE)
        return int(num.group(1)) if num else 999

    for mp in groups:
        groups[mp].sort(key=file_sort_key)

    return dict(sorted(groups.items(), key=lambda x: get_fy_sort_key(x[0])))

def compile_files(groups, log_lines, progress_bar):
    sheet_order = []
    seen        = set()

    # Discover all data sheets across all files
    for mp, files in groups.items():
        for uf in files:
            uf.seek(0)
            wb = openpyxl.load_workbook(io.BytesIO(uf.read()), read_only=True)
            for sn in wb.sheetnames:
                if sn not in SUMMARY_SHEETS and sn not in seen:
                    sheet_order.append(sn)
                    seen.add(sn)
            wb.close()

    log_lines.append(f'<span class="log-head">▸ Data sheets detected: {", ".join(sheet_order)}</span>')

    out_wb = Workbook()
    out_wb.remove(out_wb.active)

    total_months = len(groups)
    total_steps  = len(sheet_order)
    summary      = {}

    for step_idx, sheet_name in enumerate(sheet_order):
        meta_rows = get_metadata_rows(sheet_name)
        dst_ws    = out_wb.create_sheet(title=sheet_name)
        is_first  = True
        sheet_total = 0

        log_lines.append(f'<span class="log-head">━━ {sheet_name}</span>')

        for mp, files in groups.items():
            for uf in files:
                uf.seek(0)
                wb = openpyxl.load_workbook(io.BytesIO(uf.read()))
                if sheet_name not in wb.sheetnames:
                    wb.close()
                    continue
                src_ws = wb[sheet_name]
                if is_first:
                    copy_metadata(src_ws, dst_ws, meta_rows)
                    added    = append_data(src_ws, dst_ws, meta_rows)
                    is_first = False
                else:
                    added = append_data(src_ws, dst_ws, meta_rows)
                sheet_total += added
                log_lines.append(
                    f'<span class="log-dim">  [{month_label(mp)}] {uf.name.split("/")[-1]} '
                    f'→ +{added:,} rows</span>'
                )
                wb.close()

        summary[sheet_name] = sheet_total
        log_lines.append(
            f'  <span style="color:#ffffff">Total: {dst_ws.max_row:,} rows '
            f'({sheet_total:,} data + {meta_rows} header)</span>'
        )
        progress_bar.progress((step_idx + 1) / total_steps)

    # Save to buffer
    buf = io.BytesIO()
    out_wb.save(buf)
    buf.seek(0)
    return buf, summary

# ─────────────────────────────────────────────
#  UI
# ─────────────────────────────────────────────

# Header
st.markdown("""
<div class="header-banner">
    <div class="header-badge">GST Compliance Tool</div>
    <h1>📊 GSTR-2B Compiler</h1>
    <p>Upload GSTR-2B Excel files across multiple months. Get one consolidated Excel file — instantly.</p>
</div>
""", unsafe_allow_html=True)

# Info cards
st.markdown("""
<div class="info-grid">
    <div class="info-card">
        <div class="num">∞</div>
        <div class="label">Months at once</div>
    </div>
    <div class="info-card">
        <div class="num">12+</div>
        <div class="label">Sheet types handled</div>
    </div>
    <div class="info-card">
        <div class="num">100%</div>
        <div class="label">Format preserved</div>
    </div>
</div>
""", unsafe_allow_html=True)

# Upload section
st.markdown('<div class="section-label">Step 1 — Upload your files</div>', unsafe_allow_html=True)
uploaded = st.file_uploader(
    label="Drop all your GSTR-2B Excel files here",
    type=["xlsx"],
    accept_multiple_files=True,
    help="You can upload files from multiple months at once. Both single-file and multi-part months are supported."
)

if uploaded:
    groups = group_files(uploaded)
    months = list(groups.keys())

    # Show detected months
    st.markdown('<div class="section-label" style="margin-top:1.2rem">Detected months</div>', unsafe_allow_html=True)
    tags = "".join([f'<span class="month-tag">{month_label(m)}</span>' for m in months])
    st.markdown(f'<div style="margin-bottom:1rem">{tags}</div>', unsafe_allow_html=True)

    # Show file list
    st.markdown('<div class="section-label">Files loaded</div>', unsafe_allow_html=True)
    items = "".join([
        f'<div class="file-item"><div class="file-dot"></div>{uf.name}</div>'
        for mp in months for uf in groups[mp]
    ])
    st.markdown(f'<div class="file-list">{items}</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-label" style="margin-top:1.2rem">Step 2 — Compile</div>', unsafe_allow_html=True)

    if st.button("▶  Compile All Files"):
        log_lines   = []
        progress    = st.progress(0)
        log_holder  = st.empty()

        def refresh_log():
            log_html = "<br>".join(log_lines[-30:])  # show last 30 lines
            log_holder.markdown(f'<div class="log-box">{log_html}</div>', unsafe_allow_html=True)

        log_lines.append(f'<span class="log-head">Starting compilation — {len(months)} month(s)...</span>')
        refresh_log()

        try:
            buf, summary = compile_files(groups, log_lines, progress)
            refresh_log()

            log_lines.append('<span style="color:#4dca8a;font-weight:600">✓ Compilation complete!</span>')
            refresh_log()

            # Summary
            rows_html = "".join([
                f'<div class="summary-row"><span>{sn}</span>'
                f'<span class="summary-val">{cnt:,} rows</span></div>'
                for sn, cnt in summary.items()
            ])
            st.markdown(f"""
            <div class="summary-box">
                <h4>✅ Compilation Summary</h4>
                {rows_html}
            </div>
            """, unsafe_allow_html=True)

            # Determine FY label for filename
            if months:
                first_fy = get_fy_sort_key(months[0])[0]
                fy_label = f"FY{first_fy}-{str(first_fy+1)[2:]}"
            else:
                fy_label = "Compiled"

            st.markdown('<div class="section-label" style="margin-top:0.5rem">Step 3 — Download</div>', unsafe_allow_html=True)
            st.download_button(
                label="⬇  Download GSTR2B_Compiled.xlsx",
                data=buf,
                file_name=f"GSTR2B_{fy_label}_Compiled.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Error during compilation: {e}")
            log_lines.append(f'<span class="log-warn">ERROR: {e}</span>')
            refresh_log()

else:
    st.markdown("""
    <div style="background:#161b27;border:1px solid #1e2d42;border-radius:8px;
                padding:1.5rem;text-align:center;color:#3a5070;font-size:0.88rem;">
        Upload your GSTR-2B files above to get started
    </div>
    """, unsafe_allow_html=True)

# Footer
st.markdown("""
<div style="margin-top:3rem;padding-top:1rem;border-top:1px solid #1e2d42;
            text-align:center;font-size:0.72rem;color:#2a3d55;
            font-family:'IBM Plex Mono',monospace;">
    GSTR-2B Compiler • Data processed locally • Nothing uploaded to any server
</div>
""", unsafe_allow_html=True)
