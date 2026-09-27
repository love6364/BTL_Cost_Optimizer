# import streamlit as st
# import pandas as pd
# import openpyxl
# from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
# from openpyxl.utils import get_column_letter
# import io
# import re

# st.set_page_config(page_title="Diamond Data Formatter", page_icon="💎", layout="wide")

# # ── Styling ──────────────────────────────────────────────────────────────────
# st.markdown("""
# <style>
#     .main-title { font-size: 2rem; font-weight: 700; color: #1a1a2e; margin-bottom: 0; }
#     .subtitle   { color: #6c757d; margin-bottom: 1.5rem; }
#     .step-badge {
#         background: #4f46e5; color: white; border-radius: 50%;
#         width: 28px; height: 28px; display: inline-flex;
#         align-items: center; justify-content: center;
#         font-weight: 700; font-size: .85rem; margin-right: 8px;
#     }
#     .step-header { font-size: 1.1rem; font-weight: 600; color: #1a1a2e; }
#     .info-box {
#         background: #f0f4ff; border-left: 4px solid #4f46e5;
#         padding: .75rem 1rem; border-radius: 4px; margin: .5rem 0;
#     }
#     .success-box {
#         background: #f0fdf4; border-left: 4px solid #22c55e;
#         padding: .75rem 1rem; border-radius: 4px; margin: .5rem 0;
#     }
# </style>
# """, unsafe_allow_html=True)

# st.markdown('<p class="main-title">💎 Diamond Data Formatter</p>', unsafe_allow_html=True)
# st.markdown('<p class="subtitle">Upload party file → Map columns → Download formatted Excel with formulas</p>', unsafe_allow_html=True)

# # ── Constants ─────────────────────────────────────────────────────────────────
# REQUIRED_COLS = [
#     ("pexkt_no",    "Pexkt No. / Stock No."),
#     ("certi_no",    "Certi No."),
#     ("shape",       "Shape"),
#     ("carat",       "Carat"),
#     ("color",       "Color"),
#     ("clarity",     "Clarity / Clarity Final"),
#     ("final",       "Final"),
#     ("vdb",         "VDB"),
#     ("usa",         "USA"),
#     ("with_tariff", "With Tariff"),
# ]

# OUTPUT_HEADERS = [
#     "Stock No.", "Certi No.", "Shape", "Carat", "Color", "Clarity",
#     "Final", "Final Amt.", "VDB", "VDB Amt.", "USA", "USA Amt.",
#     "With Tariff", "Tariff Amt.", "Adj Cost", "Adj Amount",
#     "Profit $ Without Tariff", "Profit % Without Tariff",
#     "Adj Cost With Tariff", "Adj Amt. With Tariff",
#     "Profit $ With Tariff", "Profit % With Tariff",
# ]

# # ── Helpers ───────────────────────────────────────────────────────────────────
# def read_uploaded_file(f) -> pd.DataFrame:
#     name = f.name.lower()
#     if name.endswith(".csv"):
#         return pd.read_csv(f)
#     elif name.endswith((".xlsx", ".xls", ".xlsm")):
#         return pd.read_excel(f, header=None)
#     elif name.endswith(".tsv"):
#         return pd.read_csv(f, sep="\t")
#     else:
#         st.error("Unsupported file type. Please upload CSV, XLSX, XLS, or TSV.")
#         return None


# def detect_header_row(df: pd.DataFrame) -> int:
#     """Try to detect which row is the actual header."""
#     for i, row in df.iterrows():
#         vals = [str(v).strip().lower() for v in row if pd.notna(v)]
#         # If row has ≥4 non-null string values it's likely the header
#         if len(vals) >= 4 and not all(v.replace('.', '').replace('-', '').isdigit() for v in vals):
#             return i
#     return 0


# def apply_cell_styles(ws, header_row, data_start, data_end, n_data_cols=22):
#     """Apply professional formatting to the worksheet."""
#     header_fill   = PatternFill("solid", fgColor="1F4E79")
#     formula_fill  = PatternFill("solid", fgColor="E8F4FD")
#     summary_fill  = PatternFill("solid", fgColor="FFF2CC")
#     avg_fill      = PatternFill("solid", fgColor="E2EFDA")
#     side          = Side(style="thin", color="CCCCCC")
#     border        = Border(left=side, right=side, top=side, bottom=side)

#     for col in range(1, n_data_cols + 1):
#         cell = ws.cell(row=header_row, column=col)
#         cell.font      = Font(name="Arial", bold=True, color="FFFFFF", size=10)
#         cell.fill      = header_fill
#         cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
#         cell.border    = border

#     for row in range(data_start, data_end + 1):
#         for col in range(1, n_data_cols + 1):
#             cell = ws.cell(row=row, column=col)
#             cell.font   = Font(name="Arial", size=10)
#             cell.border = border
#             cell.alignment = Alignment(horizontal="center")
#             # Shade formula columns (H, J, L, N, O, P, Q, R, S, T, U, V)
#             if col in (8, 10, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22):
#                 cell.fill = formula_fill

#     # Summary rows
#     sum_row = data_end + 1
#     avg_row = data_end + 2
#     for col in range(1, n_data_cols + 1):
#         sc = ws.cell(row=sum_row, column=col)
#         sc.font   = Font(name="Arial", bold=True, size=10)
#         sc.fill   = summary_fill
#         sc.border = border
#         sc.alignment = Alignment(horizontal="center")
#         ac = ws.cell(row=avg_row, column=col)
#         ac.font   = Font(name="Arial", bold=True, size=10)
#         ac.fill   = avg_fill
#         ac.border = border
#         ac.alignment = Alignment(horizontal="center")

#     # Column widths
#     col_widths = [14, 14, 10, 8, 7, 9, 8, 11, 8, 11, 8, 11, 12, 12, 10, 12, 22, 22, 20, 20, 20, 20]
#     for i, w in enumerate(col_widths, 1):
#         ws.column_dimensions[get_column_letter(i)].width = w

#     ws.row_dimensions[header_row].height = 35


# def build_output_excel(mapped_df: pd.DataFrame) -> bytes:
#     """Build the final Excel file with all formulas."""
#     wb = openpyxl.Workbook()
#     ws = wb.active
#     ws.title = "Diamond Data"

#     HEADER_ROW  = 1
#     DATA_START  = 2
#     n           = len(mapped_df)
#     DATA_END    = DATA_START + n - 1
#     SUM_ROW     = DATA_END + 1
#     AVG_ROW     = DATA_END + 2
#     LABEL_ROW   = DATA_END + 5   # extra labels (Final / Rate / Difference)

#     # ── Headers ──────────────────────────────────────────────────────────────
#     for c, hdr in enumerate(OUTPUT_HEADERS, 1):
#         ws.cell(row=HEADER_ROW, column=c, value=hdr)

#     # ── Data + Formulas ───────────────────────────────────────────────────────
#     # Column letters (1-indexed to letter)
#     # A=Stock, B=Certi, C=Shape, D=Carat, E=Color, F=Clarity
#     # G=Final, H=Final Amt, I=VDB, J=VDB Amt, K=USA, L=USA Amt
#     # M=WithTariff, N=Tariff Amt, O=Adj Cost, P=Adj Amt
#     # Q=Profit$ NoTariff, R=Profit% NoTariff
#     # S=AdjCostWithTariff, T=AdjAmtWithTariff
#     # U=Profit$ WithTariff, V=Profit% WithTariff

#     # Ranges for SUMPRODUCT (USA col=K=col11, Carat col=D=col4)
#     k_range = f"$K${DATA_START}:$K${DATA_END}"
#     d_range = f"$D${DATA_START}:$D${DATA_END}"
#     h_sum   = f"$H${SUM_ROW}"   # Total Final Amt used in Adj Cost formula

#     for i, (_, row_data) in enumerate(mapped_df.iterrows()):
#         r = DATA_START + i
#         # Raw data columns
#         ws.cell(row=r, column=1,  value=row_data.get("pexkt_no"))
#         ws.cell(row=r, column=2,  value=row_data.get("certi_no"))
#         ws.cell(row=r, column=3,  value=row_data.get("shape"))
#         ws.cell(row=r, column=4,  value=row_data.get("carat"))
#         ws.cell(row=r, column=5,  value=row_data.get("color"))
#         ws.cell(row=r, column=6,  value=row_data.get("clarity"))
#         ws.cell(row=r, column=7,  value=row_data.get("final"))
#         ws.cell(row=r, column=9,  value=row_data.get("vdb"))
#         ws.cell(row=r, column=11, value=row_data.get("usa"))
#         ws.cell(row=r, column=13, value=row_data.get("with_tariff"))

#         # Formula columns
#         ws.cell(row=r, column=8,  value=f"=G{r}*D{r}")                                           # Final Amt
#         ws.cell(row=r, column=10, value=f"=I{r}*D{r}")                                           # VDB Amt
#         ws.cell(row=r, column=12, value=f"=K{r}*D{r}")                                           # USA Amt
#         ws.cell(row=r, column=14, value=f"=M{r}*D{r}")                                           # Tariff Amt
#         ws.cell(row=r, column=15, value=f"=ROUND((K{r}*{h_sum})/SUMPRODUCT({k_range},{d_range}),2)")  # Adj Cost
#         ws.cell(row=r, column=16, value=f"=O{r}*D{r}")                                           # Adj Amount
#         ws.cell(row=r, column=17, value=f"=K{r}-O{r}")                                           # Profit$ no tariff
#         ws.cell(row=r, column=18, value=f"=(Q{r}*100)/K{r}")                                     # Profit% no tariff
#         ws.cell(row=r, column=19, value=f"=O{r}*1.08")                                           # Adj Cost w Tariff
#         ws.cell(row=r, column=20, value=f"=S{r}*D{r}")                                           # Adj Amt w Tariff
#         ws.cell(row=r, column=21, value=f"=K{r}-S{r}")                                           # Profit$ w Tariff
#         ws.cell(row=r, column=22, value=f"=(U{r}*100)/O{r}")                                     # Profit% w Tariff

#         # Number formats for data cells
#         for col in (4, 7, 9, 11, 13):
#             ws.cell(row=r, column=col).number_format = "0.00"
#         for col in (8, 10, 12, 14, 15, 16, 17, 18, 19, 20, 21, 22):
#             ws.cell(row=r, column=col).number_format = "0.00"

#     # ── Summary Row (SUM) ────────────────────────────────────────────────────
#     ws.cell(row=SUM_ROW, column=1, value="TOTAL")
#     sum_cols = {4: "D", 8: "H", 10: "J", 12: "L", 14: "N", 16: "P", 20: "T"}
#     for col, letter in sum_cols.items():
#         ws.cell(row=SUM_ROW, column=col, value=f"=SUM({letter}{DATA_START}:{letter}{DATA_END})")
#         ws.cell(row=SUM_ROW, column=col).number_format = "0.00"

#     # ── Average Row ──────────────────────────────────────────────────────────
#     ws.cell(row=AVG_ROW, column=1, value="AVG / CARAT")
#     d_sum = f"D{SUM_ROW}"
#     avg_map = {8: "H", 10: "J", 12: "L", 14: "N", 16: "P", 20: "T"}
#     for col, letter in avg_map.items():
#         ws.cell(row=AVG_ROW, column=col, value=f"={letter}{SUM_ROW}/{d_sum}")
#         ws.cell(row=AVG_ROW, column=col).number_format = "0.00"

#     # ── Extra Label Section ───────────────────────────────────────────────────
#     ws.cell(row=LABEL_ROW,     column=7, value="FINAL")
#     ws.cell(row=LABEL_ROW + 1, column=7, value="$ RATE")
#     ws.cell(row=LABEL_ROW + 2, column=7, value="Difference")

#     ws.cell(row=LABEL_ROW,     column=8, value=f"=H{AVG_ROW}")
#     ws.cell(row=LABEL_ROW + 1, column=8, value="")   # user fills rate manually
#     ws.cell(row=LABEL_ROW + 2, column=8, value=f"=L{AVG_ROW}-N{AVG_ROW}")

#     for r2 in (LABEL_ROW, LABEL_ROW + 1, LABEL_ROW + 2):
#         for c2 in (7, 8):
#             cell = ws.cell(row=r2, column=c2)
#             cell.font   = Font(name="Arial", bold=True, size=10)
#             cell.border = Border(
#                 left=Side(style="thin"), right=Side(style="thin"),
#                 top=Side(style="thin"), bottom=Side(style="thin")
#             )
#             cell.alignment = Alignment(horizontal="center")
#         ws.cell(row=r2, column=8).number_format = "0.00"

#     # Rate cell highlighted for user input
#     rate_cell = ws.cell(row=LABEL_ROW + 1, column=8)
#     rate_cell.fill = PatternFill("solid", fgColor="FFFF00")
#     rate_cell.font = Font(name="Arial", bold=True, color="FF0000", size=10)

#     # ── Styles ────────────────────────────────────────────────────────────────
#     apply_cell_styles(ws, HEADER_ROW, DATA_START, DATA_END)

#     # ── Freeze panes ─────────────────────────────────────────────────────────
#     ws.freeze_panes = "A2"

#     buf = io.BytesIO()
#     wb.save(buf)
#     buf.seek(0)
#     return buf.getvalue()


# # ══════════════════════════════════════════════════════════════════════════════
# #  STEP 1 – Upload party file
# # ══════════════════════════════════════════════════════════════════════════════
# st.markdown("---")
# col_s1, _ = st.columns([1, 2])
# with col_s1:
#     st.markdown('<span class="step-badge">1</span><span class="step-header">Upload Party File</span>', unsafe_allow_html=True)

# party_file = st.file_uploader(
#     "Upload the party data file (CSV, XLSX, XLS, XLSM, TSV)",
#     type=["csv", "xlsx", "xls", "xlsm", "tsv"],
#     key="party_file"
# )

# if party_file is None:
#     st.markdown('<div class="info-box">👆 Upload the party file to get started.</div>', unsafe_allow_html=True)
#     st.stop()

# # ── Read the file ─────────────────────────────────────────────────────────────
# raw_df = read_uploaded_file(party_file)
# if raw_df is None:
#     st.stop()

# # Auto-detect header row
# h_row = detect_header_row(raw_df)
# df = raw_df.copy()
# df.columns = raw_df.iloc[h_row]
# df = df.iloc[h_row + 1:].reset_index(drop=True)
# df.columns = [str(c).strip() if pd.notna(c) else f"Col_{i}" for i, c in enumerate(df.columns)]
# df = df.dropna(how="all")

# st.success(f"✅ File loaded — **{len(df)} rows**, **{len(df.columns)} columns** detected.")

# with st.expander("👀 Preview raw data (first 5 rows)"):
#     st.dataframe(df.head(5), use_container_width=True)

# # ══════════════════════════════════════════════════════════════════════════════
# #  STEP 2 – Column Mapping
# # ══════════════════════════════════════════════════════════════════════════════
# st.markdown("---")
# st.markdown('<span class="step-badge">2</span><span class="step-header">Map Columns</span>', unsafe_allow_html=True)
# st.markdown("Select which column in the party file corresponds to each required field:")

# available_cols = ["— skip / not available —"] + list(df.columns)

# def smart_default(key_hint: str, cols: list) -> int:
#     """Try to auto-match a column by fuzzy keyword search."""
#     hints = {
#         "pexkt_no":    ["stock", "pexkt", "packet", "pkt", "id", "lot"],
#         "certi_no":    ["certi", "cert", "gia", "igi", "report"],
#         "shape":       ["shape"],
#         "carat":       ["carat", "ct", "weight"],
#         "color":       ["color", "colour"],
#         "clarity":     ["clarity", "clar", "final"],
#         "final":       ["final", "rap%", "disc%", "back%"],
#         "vdb":         ["vdb"],
#         "usa":         ["usa"],
#         "with_tariff": ["tariff", "tarrif", "with tariff", "wtariff"],
#     }
#     keywords = hints.get(key_hint, [])
#     lower_cols = [c.lower() for c in cols[1:]]  # skip "skip" entry
#     for kw in keywords:
#         for i, lc in enumerate(lower_cols):
#             if kw in lc:
#                 return i + 1  # offset by 1 for the "skip" entry
#     return 0

# mapping = {}
# col1, col2 = st.columns(2)
# for idx, (key, label) in enumerate(REQUIRED_COLS):
#     default = smart_default(key, available_cols)
#     container = col1 if idx % 2 == 0 else col2
#     with container:
#         choice = st.selectbox(
#             f"**{label}**",
#             options=available_cols,
#             index=default,
#             key=f"map_{key}"
#         )
#         mapping[key] = None if choice.startswith("—") else choice

# # Validate mandatory cols
# mandatory = ["carat", "final", "vdb", "usa", "with_tariff"]
# missing_mandatory = [k for k in mandatory if mapping.get(k) is None]
# if missing_mandatory:
#     labels = [next(l for kk, l in REQUIRED_COLS if kk == k) for k in missing_mandatory]
#     st.warning(f"⚠️ Please map these required columns: **{', '.join(labels)}**")
#     st.stop()

# # ══════════════════════════════════════════════════════════════════════════════
# #  STEP 3 – Preview & Generate
# # ══════════════════════════════════════════════════════════════════════════════
# st.markdown("---")
# st.markdown('<span class="step-badge">3</span><span class="step-header">Preview & Download</span>', unsafe_allow_html=True)

# # Build mapped dataframe
# mapped_rows = []
# for _, row in df.iterrows():
#     rec = {}
#     for key, col in mapping.items():
#         rec[key] = row[col] if col else None
#     mapped_rows.append(rec)

# mapped_df = pd.DataFrame(mapped_rows)

# # Coerce numeric cols
# for col in ["carat", "final", "vdb", "usa", "with_tariff"]:
#     mapped_df[col] = pd.to_numeric(mapped_df[col], errors="coerce")

# # Drop rows where all numeric cols are null
# mapped_df = mapped_df.dropna(subset=["carat", "final", "vdb", "usa", "with_tariff"], how="all")

# st.markdown(f"**{len(mapped_df)} rows** will be written to the output file.")

# # Preview table
# preview_cols = {
#     "pexkt_no": "Stock No.", "certi_no": "Certi No.", "shape": "Shape",
#     "carat": "Carat", "color": "Color", "clarity": "Clarity",
#     "final": "Final", "vdb": "VDB", "usa": "USA", "with_tariff": "With Tariff"
# }
# preview_df = mapped_df.rename(columns=preview_cols)
# st.dataframe(preview_df.head(10), use_container_width=True)

# if len(mapped_df) == 0:
#     st.error("No valid rows found after mapping. Please check your column selections.")
#     st.stop()

# # ── Generate Excel ────────────────────────────────────────────────────────────
# with st.spinner("Building Excel file with formulas…"):
#     xlsx_bytes = build_output_excel(mapped_df)

# st.markdown('<div class="success-box">✅ Excel file ready! Formulas are embedded — values will recalculate when you open in Excel.</div>', unsafe_allow_html=True)

# st.download_button(
#     label="⬇️ Download Formatted Excel",
#     data=xlsx_bytes,
#     file_name="Diamond_Formatted.xlsx",
#     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#     type="primary",
#     use_container_width=True,
# )

# # ── Formula Legend ────────────────────────────────────────────────────────────
# with st.expander("📐 Formula Reference"):
#     st.markdown("""
# | Column | Formula Logic |
# |---|---|
# | **Final Amt.** | `Final × Carat` |
# | **VDB Amt.** | `VDB × Carat` |
# | **USA Amt.** | `USA × Carat` |
# | **Tariff Amt.** | `With Tariff × Carat` |
# | **Adj Cost** | `ROUND((USA × Total Final Amt.) / SUMPRODUCT(USA range, Carat range), 2)` |
# | **Adj Amount** | `Adj Cost × Carat` |
# | **Profit $ (No Tariff)** | `USA − Adj Cost` |
# | **Profit % (No Tariff)** | `(Profit$ × 100) / USA` |
# | **Adj Cost With Tariff** | `Adj Cost × 1.08` |
# | **Adj Amt. With Tariff** | `Adj Cost With Tariff × Carat` |
# | **Profit $ (With Tariff)** | `USA − Adj Cost With Tariff` |
# | **Profit % (With Tariff)** | `(Profit$ With Tariff × 100) / Adj Cost` |
# | **SUM row** | Sum of Carat, Final Amt, VDB Amt, USA Amt, Tariff Amt, Adj Amt, Tariff Amt W/ Tariff |
# | **AVG row** | Each SUM ÷ Total Carat |
# | **Difference** | `USA Avg − Tariff Avg` |

# > 🟡 The **$ Rate** cell (highlighted in yellow) is left blank for you to fill in manually.
# """)




# import streamlit as st
# import pandas as pd
# import openpyxl
# from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
# from openpyxl.utils import get_column_letter
# import io
# import re
# import os

# # ─────────────────────────────────────────────────────────────────────────────
# #  PAGE CONFIG
# # ─────────────────────────────────────────────────────────────────────────────
# st.set_page_config(
#     page_title="💎 Diamond Formatter",
#     page_icon="💎",
#     layout="wide",
#     initial_sidebar_state="collapsed",
# )

# # ─────────────────────────────────────────────────────────────────────────────
# #  DARK THEME CSS
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <style>
# /* ── Base dark background ── */
# html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"] {
#     background-color: #0d0f14 !important;
#     color: #e2e8f0 !important;
# }
# [data-testid="stHeader"] { background: #0d0f14 !important; }
# [data-testid="stSidebar"] { background: #111318 !important; }
# section[data-testid="stMain"] > div { background: #0d0f14 !important; }

# /* ── Remove default padding top ── */
# .block-container { padding-top: 1.5rem !important; }

# /* ── Hero banner ── */
# .hero {
#     background: linear-gradient(135deg, #0f1629 0%, #1a0a2e 50%, #0a1628 100%);
#     border: 1px solid #2a2d3e;
#     border-radius: 16px;
#     padding: 2.5rem 2rem;
#     margin-bottom: 2rem;
#     position: relative;
#     overflow: hidden;
# }
# .hero::before {
#     content: '';
#     position: absolute; inset: 0;
#     background: radial-gradient(ellipse at 30% 50%, rgba(99,102,241,.15) 0%, transparent 60%),
#                 radial-gradient(ellipse at 70% 50%, rgba(168,85,247,.1) 0%, transparent 60%);
# }
# .hero-title {
#     font-size: 3.2rem; font-weight: 800; letter-spacing: -1px;
#     background: linear-gradient(135deg, #a78bfa, #6366f1, #38bdf8);
#     -webkit-background-clip: text; -webkit-text-fill-color: transparent;
#     margin: 0; position: relative;
# }
# .hero-sub {
#     color: #94a3b8; font-size: 0.50rem; margin-top: .3rem; position: relative;
# }
# .hero-badge {
#     display: inline-block; background: rgba(99,102,241,.2);
#     border: 1px solid rgba(99,102,241,.4); border-radius: 20px;
#     padding: 3px 12px; font-size: .75rem; color: #a78bfa;
#     margin-bottom: .75rem; position: relative;
# }

# /* ── Color legend pills ── */
# .legend-wrap { display: flex; flex-wrap: wrap; gap: 8px; margin: 1rem 0 1.5rem; }
# .pill {
#     display: inline-flex; align-items: center; gap: 7px;
#     background: #1a1d27; border: 1px solid #2a2d3e;
#     border-radius: 20px; padding: 5px 12px; font-size: .78rem; color: #cbd5e1;
# }
# .dot { width: 12px; height: 12px; border-radius: 3px; flex-shrink: 0; }

# /* ── Step cards ── */
# .step-card {
#     background: #13161f;
#     border: 1px solid #1e2130;
#     border-radius: 12px;
#     padding: 1.5rem;
#     margin-bottom: 1rem;
# }
# .step-num {
#     display: inline-flex; align-items: center; justify-content: center;
#     width: 32px; height: 32px; border-radius: 50%;
#     background: linear-gradient(135deg, #6366f1, #8b5cf6);
#     color: white; font-weight: 700; font-size: .9rem;
#     margin-right: 10px; flex-shrink: 0;
# }
# .step-title { font-size: 1.1rem; font-weight: 600; color: #f1f5f9; }
# .step-header { display: flex; align-items: center; margin-bottom: 1rem; }

# /* ── Streamlit widget overrides (dark) ── */
# [data-testid="stFileUploader"] {
#     background: #13161f !important;
#     border: 2px dashed #2a2d3e !important;
#     border-radius: 10px !important;
# }
# [data-testid="stFileUploader"]:hover { border-color: #6366f1 !important; }
# .stSelectbox > div > div {
#     background: #1a1d27 !important;
#     border: 1px solid #2a2d3e !important;
#     color: #e2e8f0 !important;
#     border-radius: 8px !important;
# }
# .stSelectbox label { color: #cbd5e1 !important; font-size: .85rem !important; }
# [data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }

# /* ── Download button ── */
# .stDownloadButton button {
#     background: linear-gradient(135deg, #6366f1, #8b5cf6) !important;
#     color: white !important; border: none !important;
#     border-radius: 10px !important; font-weight: 600 !important;
#     padding: .75rem 1.5rem !important; font-size: 1rem !important;
#     transition: all .2s !important; box-shadow: 0 4px 20px rgba(99,102,241,.4) !important;
# }
# .stDownloadButton button:hover {
#     transform: translateY(-2px) !important;
#     box-shadow: 0 8px 30px rgba(99,102,241,.6) !important;
# }

# /* ── Success / info / warning banners ── */
# .banner {
#     border-radius: 10px; padding: .85rem 1.1rem;
#     font-size: .9rem; margin: .5rem 0; display: flex; align-items: center; gap: 10px;
# }
# .banner-success { background: rgba(34,197,94,.1); border: 1px solid rgba(34,197,94,.3); color: #86efac; }
# .banner-info    { background: rgba(99,102,241,.1); border: 1px solid rgba(99,102,241,.3); color: #a5b4fc; }
# .banner-warn    { background: rgba(234,179,8,.1);  border: 1px solid rgba(234,179,8,.3);  color: #fde68a; }

# /* ── Stats row ── */
# .stats-row { display: flex; gap: 12px; flex-wrap: wrap; margin: 1rem 0; }
# .stat-box {
#     background: #13161f; border: 1px solid #1e2130; border-radius: 10px;
#     padding: .75rem 1.25rem; flex: 1; min-width: 130px;
# }
# .stat-val { font-size: 1.6rem; font-weight: 700; color: #a78bfa; }
# .stat-label { font-size: .75rem; color: #64748b; text-transform: uppercase; letter-spacing: .5px; }

# /* ── Divider ── */
# hr { border-color: #1e2130 !important; margin: 1.5rem 0 !important; }

# /* ── Expander dark ── */
# [data-testid="stExpander"] {
#     background: #13161f !important; border: 1px solid #1e2130 !important;
#     border-radius: 10px !important;
# }
# summary { color: #94a3b8 !important; }

# /* ── Dataframe dark cells ── */
# .stDataFrame thead th { background: #1a1d27 !important; color: #a78bfa !important; }
# .stDataFrame tbody tr:nth-child(even) td { background: #13161f !important; }
# .stDataFrame tbody tr:nth-child(odd)  td { background: #0f111a !important; }
# </style>
# """, unsafe_allow_html=True)

# # ─────────────────────────────────────────────────────────────────────────────
# #  EXACT EXCEL COLORS (from Diamond_Formatted.xlsx)
# # ─────────────────────────────────────────────────────────────────────────────
# CLR_HEADER   = "1F4E79"   # deep navy  → col A-V header row
# CLR_CARAT    = "98F6EB"   # teal mint  → D (Carat)
# CLR_FINAL    = "FFCF37"   # golden     → G, H (Final, Final Amt)
# CLR_VDB      = "5991D5"   # steel blue → I, J (VDB, VDB Amt)
# CLR_USA      = "F79B4F"   # orange     → K, L (USA, USA Amt)
# CLR_TARIFF   = "9E89B9"   # lavender   → M, N (With Tariff, Tariff Amt)
# CLR_ADJ      = "98B850"   # sage green → O, P, Q, R (Adj Cost group)
# CLR_WTARIFF  = "CB6967"   # coral red  → S, T, U, V (With Tariff group)

# # col index → fill color (1-based)
# COL_COLORS = {
#     4:  CLR_CARAT,
#     7:  CLR_FINAL,   8:  CLR_FINAL,
#     9:  CLR_VDB,     10: CLR_VDB,
#     11: CLR_USA,     12: CLR_USA,
#     13: CLR_TARIFF,  14: CLR_TARIFF,
#     15: CLR_ADJ,     16: CLR_ADJ,  17: CLR_ADJ,  18: CLR_ADJ,
#     19: CLR_WTARIFF, 20: CLR_WTARIFF, 21: CLR_WTARIFF, 22: CLR_WTARIFF,
# }

# OUTPUT_HEADERS = [
#     "Stock No.", "Certi No.", "Shape", "Carat", "Color", "Clarity",
#     "Final", "Final Amt.", "VDB", "VDB Amt.", "USA", "USA Amt.",
#     "With Tariff", "Tariff Amt.", "Adj Cost", "Adj Amount",
#     "Profit $ Without Tariff", "Profit % Without Tariff",
#     "Adj Cost With Tariff", "Adj Amt. With Tariff",
#     "Profit $ With Tariff", "Profit % With Tariff",
# ]

# REQUIRED_COLS = [
#     ("pexkt_no",    "Pexkt No. / Stock No."),
#     ("certi_no",    "Certi No."),
#     ("shape",       "Shape"),
#     ("carat",       "Carat"),
#     ("color",       "Color"),
#     ("clarity",     "Clarity"),
#     ("final",       "Final"),
#     ("vdb",         "VDB"),
#     ("usa",         "USA"),
#     ("with_tariff", "With Tariff"),
# ]

# # ─────────────────────────────────────────────────────────────────────────────
# #  HELPERS
# # ─────────────────────────────────────────────────────────────────────────────
# def read_uploaded_file(f) -> pd.DataFrame | None:
#     name = f.name.lower()
#     try:
#         if name.endswith(".csv"):
#             return pd.read_csv(f)
#         elif name.endswith((".xlsx", ".xls", ".xlsm")):
#             return pd.read_excel(f, header=None)
#         elif name.endswith(".tsv"):
#             return pd.read_csv(f, sep="\t")
#         else:
#             st.error("Unsupported format. Please upload CSV, XLSX, XLS, XLSM, or TSV.")
#             return None
#     except Exception as e:
#         st.error(f"Could not read file: {e}")
#         return None


# def detect_header_row(df: pd.DataFrame) -> int:
#     for i, row in df.iterrows():
#         vals = [str(v).strip() for v in row if pd.notna(v) and str(v).strip()]
#         if len(vals) >= 4 and not all(
#             re.sub(r'[.\-]', '', v).isdigit() for v in vals
#         ):
#             return i
#     return 0


# def safe_filename(raw_name: str, n_stones: int) -> str:
#     """PartyFile_19_Stones.xlsx"""
#     stem = os.path.splitext(raw_name)[0]
#     stem = re.sub(r'[^\w\s\-]', '', stem).strip()
#     stem = re.sub(r'\s+', '_', stem)
#     return f"{stem}_{n_stones}_Stones.xlsx"


# def smart_default(key: str, cols: list) -> int:
#     hints = {
#         "pexkt_no":    ["stock", "pexkt", "packet", "pkt", "id", "lot"],
#         "certi_no":    ["certi", "cert", "gia", "igi", "report"],
#         "shape":       ["shape"],
#         "carat":       ["carat", "ct", "weight"],
#         "color":       ["color", "colour"],
#         "clarity":     ["clarity", "clar"],
#         "final":       ["final", "rap%", "disc%", "back%"],
#         "vdb":         ["vdb"],
#         "usa":         ["usa"],
#         "with_tariff": ["tariff", "tarrif", "with tariff", "wtariff"],
#     }
#     keywords = hints.get(key, [])
#     lower_cols = [c.lower() for c in cols[1:]]
#     for kw in keywords:
#         for i, lc in enumerate(lower_cols):
#             if kw in lc:
#                 return i + 1
#     return 0


# def px(hex_color: str) -> PatternFill:
#     return PatternFill("solid", fgColor=hex_color)


# def thin_border() -> Border:
#     s = Side(style="thin", color="D0D0D0")
#     return Border(left=s, right=s, top=s, bottom=s)


# def build_output_excel(mapped_df: pd.DataFrame) -> bytes:
#     wb = openpyxl.Workbook()
#     ws = wb.active
#     ws.title = "Diamond Data"

#     n          = len(mapped_df)
#     HDR        = 1
#     DS         = 2          # data start
#     DE         = DS + n - 1 # data end
#     SUM_ROW    = DE + 1
#     AVG_ROW    = DE + 2
#     LBL_ROW    = DE + 5     # FINAL / RATE / Difference

#     k_rng = f"$K${DS}:$K${DE}"
#     d_rng = f"$D${DS}:$D${DE}"
#     h_sum = f"$H${SUM_ROW}"

#     # ── Header row ──────────────────────────────────────────────────────────
#     for c, hdr in enumerate(OUTPUT_HEADERS, 1):
#         cell = ws.cell(row=HDR, column=c, value=hdr)
#         cell.fill      = px(CLR_HEADER)
#         cell.font      = Font(name="Arial", bold=True, color="FFFFFF", size=10)
#         cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
#         cell.border    = thin_border()
#     ws.row_dimensions[HDR].height = 36

#     # ── Data rows ────────────────────────────────────────────────────────────
#     for i, (_, row) in enumerate(mapped_df.iterrows()):
#         r = DS + i

#         # Raw value columns
#         raw = {
#             1: row.get("pexkt_no"),
#             2: row.get("certi_no"),
#             3: row.get("shape"),
#             4: row.get("carat"),
#             5: row.get("color"),
#             6: row.get("clarity"),
#             7: row.get("final"),
#             9: row.get("vdb"),
#             11: row.get("usa"),
#             13: row.get("with_tariff"),
#         }
#         # Formula columns
#         formulas = {
#             8:  f"=G{r}*D{r}",
#             10: f"=I{r}*D{r}",
#             12: f"=K{r}*D{r}",
#             14: f"=M{r}*D{r}",
#             15: f"=ROUND((K{r}*{h_sum})/SUMPRODUCT({k_rng},{d_rng}),2)",
#             16: f"=O{r}*D{r}",
#             17: f"=K{r}-O{r}",
#             18: f"=(Q{r}*100)/K{r}",
#             19: f"=O{r}*1.08",
#             20: f"=S{r}*D{r}",
#             21: f"=K{r}-S{r}",
#             22: f"=(U{r}*100)/O{r}",
#         }

#         for c in range(1, 23):
#             val = raw.get(c, formulas.get(c))
#             cell = ws.cell(row=r, column=c, value=val)
#             cell.border    = thin_border()
#             cell.alignment = Alignment(horizontal="center", vertical="center")
#             cell.font      = Font(name="Arial", size=10)
#             if c in COL_COLORS:
#                 cell.fill = px(COL_COLORS[c])
#             if c in (4, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22):
#                 cell.number_format = "0.00"

#         ws.row_dimensions[r].height = 18

#     # ── TOTAL row ─────────────────────────────────────────────────────────
#     total_fill = px("F2F2F2")
#     ws.cell(row=SUM_ROW, column=1, value="TOTAL").font = Font(name="Arial", bold=True, size=10)
#     for col, letter in {4:"D", 8:"H", 10:"J", 12:"L", 14:"N", 16:"P", 20:"T"}.items():
#         cell = ws.cell(row=SUM_ROW, column=col,
#                        value=f"=SUM({letter}{DS}:{letter}{DE})")
#         cell.font          = Font(name="Arial", bold=True, size=10)
#         cell.fill          = total_fill
#         cell.border        = thin_border()
#         cell.alignment     = Alignment(horizontal="center")
#         cell.number_format = "0.00"
#     ws.cell(row=SUM_ROW, column=1).fill   = total_fill
#     ws.cell(row=SUM_ROW, column=1).border = thin_border()
#     ws.cell(row=SUM_ROW, column=1).alignment = Alignment(horizontal="center")

#     # ── AVG row ───────────────────────────────────────────────────────────
#     avg_fill = px("E2EFDA")
#     ws.cell(row=AVG_ROW, column=1, value="AVG / CARAT").font = Font(name="Arial", bold=True, size=10)
#     for col, letter in {8:"H", 10:"J", 12:"L", 14:"N", 16:"P", 20:"T"}.items():
#         cell = ws.cell(row=AVG_ROW, column=col,
#                        value=f"={letter}{SUM_ROW}/D{SUM_ROW}")
#         cell.font          = Font(name="Arial", bold=True, size=10)
#         cell.fill          = avg_fill
#         cell.border        = thin_border()
#         cell.alignment     = Alignment(horizontal="center")
#         cell.number_format = "0.00"
#     ws.cell(row=AVG_ROW, column=1).fill   = avg_fill
#     ws.cell(row=AVG_ROW, column=1).border = thin_border()
#     ws.cell(row=AVG_ROW, column=1).alignment = Alignment(horizontal="center")

#     # ── Label block (FINAL / $ RATE / Difference) ─────────────────────────
#     lbl_data = [
#         (LBL_ROW,     "FINAL",      f"=H{AVG_ROW}"),
#         (LBL_ROW + 1, "$ RATE",     ""),
#         (LBL_ROW + 2, "Difference", f"=L{AVG_ROW}-N{AVG_ROW}"),
#     ]
#     for lrow, lbl, formula in lbl_data:
#         for c, val in ((7, lbl), (8, formula)):
#             cell = ws.cell(row=lrow, column=c, value=val)
#             cell.font      = Font(name="Arial", bold=True, size=10)
#             cell.border    = thin_border()
#             cell.alignment = Alignment(horizontal="center")
#             cell.number_format = "0.00"
#     # Highlight $ RATE cell for manual entry
#     rate_cell = ws.cell(row=LBL_ROW + 1, column=8)
#     rate_cell.fill = px("FFFF00")
#     rate_cell.font = Font(name="Arial", bold=True, color="FF0000", size=10)

#     # ── Column widths ──────────────────────────────────────────────────────
#     widths = [14, 14, 11, 8, 7, 9, 8, 11, 8, 11, 8, 11, 12, 12, 10, 12, 24, 22, 22, 22, 22, 22]
#     for i, w in enumerate(widths, 1):
#         ws.column_dimensions[get_column_letter(i)].width = w

#     ws.freeze_panes = "A2"

#     buf = io.BytesIO()
#     wb.save(buf)
#     buf.seek(0)
#     return buf.getvalue()


# # ─────────────────────────────────────────────────────────────────────────────
# #  HERO BANNER
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="hero">
#   <div class="hero-badge">✦ Bright Trading India LLP</div>
#   <p class="hero-title">💎 DIAMOND COST OPTIMIZER</p>
#   <p class="hero-sub">Smart Cost Adjustment & Equal Profit Calculator</p>
# </div>
# """, unsafe_allow_html=True)

# # Color legend
# st.markdown("""
# <div class="legend-wrap">
#   <div class="pill"><div class="dot" style="background:#98F6EB"></div>Carat</div>
#   <div class="pill"><div class="dot" style="background:#FFCF37"></div>Final / Final Amt</div>
#   <div class="pill"><div class="dot" style="background:#5991D5"></div>VDB / VDB Amt</div>
#   <div class="pill"><div class="dot" style="background:#F79B4F"></div>USA / USA Amt</div>
#   <div class="pill"><div class="dot" style="background:#9E89B9"></div>With Tariff / With Tariff Amt</div>
#   <div class="pill"><div class="dot" style="background:#98B850"></div>Adj Cost Group</div>
#   <div class="pill"><div class="dot" style="background:#CB6967"></div>Tariff Group</div>
# </div>
# """, unsafe_allow_html=True)

# # ─────────────────────────────────────────────────────────────────────────────
# #  STEP 1 – UPLOAD
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="step-card">
#   <div class="step-header">
#     <span class="step-num">1</span>
#     <span class="step-title">Upload Party File</span>
#   </div>    
# </div>
# """, unsafe_allow_html=True)

# party_file = st.file_uploader(
#     "  Start by uploading the party file (CSV, XLSX, XLS, XLSM, or TSV).",
#     type=["csv", "xlsx", "xls", "xlsm", "tsv"],
#     key="party_file",
#     label_visibility="visible",
# )

# if party_file is None:
#     st.markdown('<div class="banner banner-info">👆 Upload the party file above to begin.</div>',
#                 unsafe_allow_html=True)
#     st.stop()

# # ── Read & normalise ──────────────────────────────────────────────────────────
# raw_df = read_uploaded_file(party_file)
# if raw_df is None:
#     st.stop()

# h_row = detect_header_row(raw_df)
# df = raw_df.copy()
# df.columns = raw_df.iloc[h_row]
# df = df.iloc[h_row + 1:].reset_index(drop=True)
# df.columns = [str(c).strip() if pd.notna(c) else f"Col_{i}" for i, c in enumerate(df.columns)]
# df = df.dropna(how="all")

# # Stats
# st.markdown(f"""
# <div class="stats-row">
#   <div class="stat-box"><div class="stat-val">{len(df)}</div><div class="stat-label">Rows Detected</div></div>
#   <div class="stat-box"><div class="stat-val">{len(df.columns)}</div><div class="stat-label">Columns Found</div></div>
#   <div class="stat-box"><div class="stat-val">{party_file.name.rsplit(".",1)[-1].upper()}</div><div class="stat-label">File Format</div></div>
# </div>
# """, unsafe_allow_html=True)

# with st.expander("👀 Preview raw data (first 5 rows)"):
#     st.dataframe(df.head(5), use_container_width=True)

# # ─────────────────────────────────────────────────────────────────────────────
# #  STEP 2 – COLUMN MAPPING
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="step-card">
#   <div class="step-header">
#     <span class="step-num">2</span>
#     <span class="step-title">Map Columns</span>
#   </div>
# </div>
# """, unsafe_allow_html=True)
# st.caption("Select which column in the party file corresponds to each required field. Auto-matched where possible.")

# available = ["— skip / not available —"] + list(df.columns)
# mapping   = {}

# cols_left, cols_right = st.columns(2)
# for idx, (key, label) in enumerate(REQUIRED_COLS):
#     default = smart_default(key, available)
#     container = cols_left if idx % 2 == 0 else cols_right
#     with container:
#         choice = st.selectbox(f"**{label}**", options=available,
#                               index=default, key=f"map_{key}")
#         mapping[key] = None if choice.startswith("—") else choice

# # Validate mandatory
# mandatory = ["carat", "final", "vdb", "usa", "with_tariff"]
# missing   = [next(l for kk, l in REQUIRED_COLS if kk == k)
#              for k in mandatory if not mapping.get(k)]
# if missing:
#     st.markdown(f'<div class="banner banner-warn">⚠️ Please map required columns: <b>{", ".join(missing)}</b></div>',
#                 unsafe_allow_html=True)
#     st.stop()

# # ─────────────────────────────────────────────────────────────────────────────
# #  STEP 3 – PREVIEW & DOWNLOAD
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="step-card">
#   <div class="step-header">
#     <span class="step-num">3</span>
#     <span class="step-title">Preview & Download</span>
#   </div>
# </div>
# """, unsafe_allow_html=True)

# # Build mapped dataframe
# records = []
# for _, row in df.iterrows():
#     rec = {key: (row[col] if col else None) for key, col in mapping.items()}
#     records.append(rec)

# mapped_df = pd.DataFrame(records)
# for col in ["carat", "final", "vdb", "usa", "with_tariff"]:
#     mapped_df[col] = pd.to_numeric(mapped_df[col], errors="coerce")

# mapped_df = mapped_df.dropna(subset=["carat", "final", "vdb", "usa", "with_tariff"], how="all")
# n_stones  = len(mapped_df)

# if n_stones == 0:
#     st.markdown('<div class="banner banner-warn">⚠️ No valid rows found. Check your column mappings.</div>',
#                 unsafe_allow_html=True)
#     st.stop()

# st.markdown(f'<div class="banner banner-success">✅ <b>{n_stones} stones</b> ready to export.</div>',
#             unsafe_allow_html=True)

# preview_rename = {
#     "pexkt_no":"Stock No.", "certi_no":"Certi No.", "shape":"Shape",
#     "carat":"Carat", "color":"Color", "clarity":"Clarity",
#     "final":"Final", "vdb":"VDB", "usa":"USA", "with_tariff":"With Tariff",
# }
# st.dataframe(mapped_df.rename(columns=preview_rename).head(10), use_container_width=True)

# # ── Generate Excel ────────────────────────────────────────────────────────────
# with st.spinner("Building Excel with formulas and colours…"):
#     xlsx_bytes  = build_output_excel(mapped_df)
#     out_filename = safe_filename(party_file.name, n_stones)

# st.markdown('<div class="banner banner-success">💾 File ready — formulas & colours match your reference format exactly.</div>',
#             unsafe_allow_html=True)

# st.download_button(
#     label=f"⬇️  Download  {out_filename}",
#     data=xlsx_bytes,
#     file_name=out_filename,
#     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#     use_container_width=True,
# )

# # ── Formula reference ─────────────────────────────────────────────────────────
# with st.expander("📐 Formula & Colour Reference"):
#     st.markdown("""
# | Column | Colour | Source / Formula |
# |---|---|---|
# | Stock No., Certi No., Shape, Color, Clarity | — | From party file |
# | **Carat** | 🟦 Teal `#98F6EB` | From party file |
# | **Final**, **Final Amt** | 🟨 Gold `#FFCF37` | Final from party; Amt = `Final × Carat` |
# | **VDB**, **VDB Amt** | 🔵 Blue `#5991D5` | VDB from party; Amt = `VDB × Carat` |
# | **USA**, **USA Amt** | 🟠 Orange `#F79B4F` | USA from party; Amt = `USA × Carat` |
# | **With Tariff**, **Tariff Amt** | 🟣 Lavender `#9E89B9` | From party; Amt = `With Tariff × Carat` |
# | **Adj Cost** | 🟩 Green `#98B850` | `ROUND((USA × ΣFinalAmt) / SUMPRODUCT(USA,Carat), 2)` |
# | **Adj Amount** | 🟩 | `Adj Cost × Carat` |
# | **Profit $ (No Tariff)** | 🟩 | `USA − Adj Cost` |
# | **Profit % (No Tariff)** | 🟩 | `(Profit$ × 100) / USA` |
# | **Adj Cost With Tariff** | 🟥 Coral `#CB6967` | `Adj Cost × 1.08` |
# | **Adj Amt With Tariff** | 🟥 | `Adj Cost With Tariff × Carat` |
# | **Profit $ (With Tariff)** | 🟥 | `USA − Adj Cost With Tariff` |
# | **Profit % (With Tariff)** | 🟥 | `(Profit$ With Tariff × 100) / Adj Cost` |
# | **TOTAL row** | Light grey | SUM of each value column |
# | **AVG / CARAT row** | Light green | Each SUM ÷ Total Carat |
# | **Difference** | — | `USA Avg − Tariff Avg` |

# > 🟡 The **$ RATE** cell is highlighted yellow — fill it manually in Excel.
# """)

# import streamlit as st
# import pandas as pd
# import openpyxl
# from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
# from openpyxl.utils import get_column_letter
# import io
# import re
# import os

# # ─────────────────────────────────────────────────────────────────────────────
# #  PAGE CONFIG
# # ─────────────────────────────────────────────────────────────────────────────
# st.set_page_config(
#     page_title="💎 Diamond Formatter",
#     page_icon="💎",
#     layout="wide",
#     initial_sidebar_state="collapsed",
# )

# # ─────────────────────────────────────────────────────────────────────────────
# #  DARK THEME CSS
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <style>
# /* ── Base dark background ── */
# html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"] {
#     background-color: #0d0f14 !important;
#     color: #e2e8f0 !important;
# }
# [data-testid="stHeader"] { background: #0d0f14 !important; }
# [data-testid="stSidebar"] { background: #111318 !important; }
# section[data-testid="stMain"] > div { background: #0d0f14 !important; }

# /* ── Remove default padding top ── */
# .block-container { padding-top: 1.5rem !important; }

# /* ── Hero banner ── */
# .hero {
#     background: linear-gradient(135deg, #0f1629 0%, #1a0a2e 50%, #0a1628 100%);
#     border: 1px solid #2a2d3e;
#     border-radius: 16px;
#     padding: 2.5rem 2rem;
#     margin-bottom: 2rem;
#     position: relative;
#     overflow: hidden;
# }
# .hero::before {
#     content: '';
#     position: absolute; inset: 0;
#     background: radial-gradient(ellipse at 30% 50%, rgba(99,102,241,.15) 0%, transparent 60%),
#                 radial-gradient(ellipse at 70% 50%, rgba(168,85,247,.1) 0%, transparent 60%);
# }
# .hero-title {
#     font-size: 2.6rem; font-weight: 800; letter-spacing: -1px;
#     background: linear-gradient(135deg, #a78bfa, #6366f1, #38bdf8);
#     -webkit-background-clip: text; -webkit-text-fill-color: transparent;
#     margin: 0; position: relative;
# }
# .hero-sub {
#     color: #94a3b8; font-size: 1.05rem; margin-top: .5rem; position: relative;
# }
# .hero-badge {
#     display: inline-block; background: rgba(99,102,241,.2);
#     border: 1px solid rgba(99,102,241,.4); border-radius: 20px;
#     padding: 3px 12px; font-size: .75rem; color: #a78bfa;
#     margin-bottom: .75rem; position: relative;
# }

# /* ── Color legend pills ── */
# .legend-wrap { display: flex; flex-wrap: wrap; gap: 8px; margin: 1rem 0 1.5rem; }
# .pill {
#     display: inline-flex; align-items: center; gap: 7px;
#     background: #1a1d27; border: 1px solid #2a2d3e;
#     border-radius: 20px; padding: 5px 12px; font-size: .78rem; color: #cbd5e1;
# }
# .dot { width: 12px; height: 12px; border-radius: 3px; flex-shrink: 0; }

# /* ── Step cards ── */
# .step-card {
#     background: #13161f;
#     border: 1px solid #1e2130;
#     border-radius: 12px;
#     padding: 1.5rem;
#     margin-bottom: 1rem;
# }
# .step-num {
#     display: inline-flex; align-items: center; justify-content: center;
#     width: 32px; height: 32px; border-radius: 50%;
#     background: linear-gradient(135deg, #6366f1, #8b5cf6);
#     color: white; font-weight: 700; font-size: .9rem;
#     margin-right: 10px; flex-shrink: 0;
# }
# .step-title { font-size: 1.1rem; font-weight: 600; color: #f1f5f9; }
# .step-header { display: flex; align-items: center; margin-bottom: 1rem; }

# /* ── Streamlit widget overrides (dark) ── */
# [data-testid="stFileUploader"] {
#     background: #13161f !important;
#     border: 2px dashed #2a2d3e !important;
#     border-radius: 10px !important;
# }
# [data-testid="stFileUploader"]:hover { border-color: #6366f1 !important; }
# .stSelectbox > div > div {
#     background: #1a1d27 !important;
#     border: 1px solid #2a2d3e !important;
#     color: #e2e8f0 !important;
#     border-radius: 8px !important;
# }
# .stSelectbox label { color: #cbd5e1 !important; font-size: .85rem !important; }
# [data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }

# /* ── Download button ── */
# .stDownloadButton button {
#     background: linear-gradient(135deg, #6366f1, #8b5cf6) !important;
#     color: white !important; border: none !important;
#     border-radius: 10px !important; font-weight: 600 !important;
#     padding: .75rem 1.5rem !important; font-size: 1rem !important;
#     transition: all .2s !important; box-shadow: 0 4px 20px rgba(99,102,241,.4) !important;
# }
# .stDownloadButton button:hover {
#     transform: translateY(-2px) !important;
#     box-shadow: 0 8px 30px rgba(99,102,241,.6) !important;
# }

# /* ── Success / info / warning banners ── */
# .banner {
#     border-radius: 10px; padding: .85rem 1.1rem;
#     font-size: .9rem; margin: .5rem 0; display: flex; align-items: center; gap: 10px;
# }
# .banner-success { background: rgba(34,197,94,.1); border: 1px solid rgba(34,197,94,.3); color: #86efac; }
# .banner-info    { background: rgba(99,102,241,.1); border: 1px solid rgba(99,102,241,.3); color: #a5b4fc; }
# .banner-warn    { background: rgba(234,179,8,.1);  border: 1px solid rgba(234,179,8,.3);  color: #fde68a; }

# /* ── Stats row ── */
# .stats-row { display: flex; gap: 12px; flex-wrap: wrap; margin: 1rem 0; }
# .stat-box {
#     background: #13161f; border: 1px solid #1e2130; border-radius: 10px;
#     padding: .75rem 1.25rem; flex: 1; min-width: 130px;
# }
# .stat-val { font-size: 1.6rem; font-weight: 700; color: #a78bfa; }
# .stat-label { font-size: .75rem; color: #64748b; text-transform: uppercase; letter-spacing: .5px; }

# /* ── Divider ── */
# hr { border-color: #1e2130 !important; margin: 1.5rem 0 !important; }

# /* ── Expander dark ── */
# [data-testid="stExpander"] {
#     background: #13161f !important; border: 1px solid #1e2130 !important;
#     border-radius: 10px !important;
# }
# summary { color: #94a3b8 !important; }

# /* ── Dataframe dark cells ── */
# .stDataFrame thead th { background: #1a1d27 !important; color: #a78bfa !important; }
# .stDataFrame tbody tr:nth-child(even) td { background: #13161f !important; }
# .stDataFrame tbody tr:nth-child(odd)  td { background: #0f111a !important; }
# </style>
# """, unsafe_allow_html=True)

# # ─────────────────────────────────────────────────────────────────────────────
# #  EXACT EXCEL COLORS (from Diamond_Formatted.xlsx)
# # ─────────────────────────────────────────────────────────────────────────────
# CLR_HEADER   = "1F4E79"   # deep navy  → col A-V header row
# CLR_CARAT    = "98F6EB"   # teal mint  → D (Carat)
# CLR_FINAL    = "FFCF37"   # golden     → G, H (Final, Final Amt)
# CLR_VDB      = "5991D5"   # steel blue → I, J (VDB, VDB Amt)
# CLR_USA      = "F79B4F"   # orange     → K, L (USA, USA Amt)
# CLR_TARIFF   = "9E89B9"   # lavender   → M, N (With Tariff, Tariff Amt)
# CLR_ADJ      = "98B850"   # sage green → O, P, Q, R (Adj Cost group)
# CLR_WTARIFF  = "CB6967"   # coral red  → S, T, U, V (With Tariff group)

# # col index → fill color (1-based)
# COL_COLORS = {
#     4:  CLR_CARAT,
#     7:  CLR_FINAL,   8:  CLR_FINAL,
#     9:  CLR_VDB,     10: CLR_VDB,
#     11: CLR_USA,     12: CLR_USA,
#     13: CLR_TARIFF,  14: CLR_TARIFF,
#     15: CLR_ADJ,     16: CLR_ADJ,  17: CLR_ADJ,  18: CLR_ADJ,
#     19: CLR_WTARIFF, 20: CLR_WTARIFF, 21: CLR_WTARIFF, 22: CLR_WTARIFF,
# }

# OUTPUT_HEADERS = [
#     "Stock ID.", "Certi No.", "Shape", "Carat", "Color", "Clarity",
#     "Final", "Final Amt.", "VDB", "VDB Amt.", "USA", "USA Amt.",
#     "With Tariff", "Tariff Amt.", "Adj Cost", "Adj Amount",
#     "Profit $ Without Tariff", "Profit % Without Tariff",
#     "Adj Cost With Tariff", "Adj Amt. With Tariff",
#     "Profit $ With Tariff", "Profit % With Tariff",
# ]

# REQUIRED_COLS = [
#     ("stock_no",    "Stock ID."),
#     ("certi_no",    "Certi No."),
#     ("shape",       "Shape"),
#     ("carat",       "Carat"),
#     ("color",       "Color"),
#     ("clarity",     "Clarity"),
#     ("final",       "Final"),
#     ("vdb",         "VDB"),
#     ("usa",         "USA"),
#     ("with_tariff", "With Tariff"),
# ]

# # ─────────────────────────────────────────────────────────────────────────────
# #  HELPERS
# # ─────────────────────────────────────────────────────────────────────────────
# def read_uploaded_file(f) -> pd.DataFrame | None:
#     name = f.name.lower()
#     try:
#         if name.endswith(".csv"):
#             return pd.read_csv(f)
#         elif name.endswith((".xlsx", ".xls", ".xlsm")):
#             return pd.read_excel(f, header=None)
#         elif name.endswith(".tsv"):
#             return pd.read_csv(f, sep="\t")
#         else:
#             st.error("Unsupported format. Please upload CSV, XLSX, XLS, XLSM, or TSV.")
#             return None
#     except Exception as e:
#         st.error(f"Could not read file: {e}")
#         return None


# def detect_header_row(df: pd.DataFrame) -> int:
#     for i, row in df.iterrows():
#         vals = [str(v).strip() for v in row if pd.notna(v) and str(v).strip()]
#         if len(vals) >= 4 and not all(
#             re.sub(r'[.\-]', '', v).isdigit() for v in vals
#         ):
#             return i
#     return 0


# def safe_filename(raw_name: str, n_stones: int) -> str:
#     """PartyFile_19_Stones.xlsx"""
#     stem = os.path.splitext(raw_name)[0]
#     stem = re.sub(r'[^\w\s\-]', '', stem).strip()
#     stem = re.sub(r'\s+', '_', stem)
#     return f"{stem}_{n_stones}_Stones.xlsx"


# def smart_default(key: str, cols: list) -> int:
#     hints = {
#         "stock_no":    ["stock", "pexkt", "packet", "pkt", "id", "lot"],
#         "certi_no":    ["certi", "cert", "gia", "igi", "report"],
#         "shape":       ["shape"],
#         "carat":       ["carat", "ct", "weight"],
#         "color":       ["color", "colour"],
#         "clarity":     ["clarity", "clar"],
#         "final":       ["final", "rap%", "disc%", "back%"],
#         "vdb":         ["vdb"],
#         "usa":         ["usa"],
#         "with_tariff": ["tariff", "tarrif", "with tariff", "wtariff"],
#     }
#     keywords = hints.get(key, [])
#     lower_cols = [c.lower() for c in cols[1:]]
#     for kw in keywords:
#         for i, lc in enumerate(lower_cols):
#             if kw in lc:
#                 return i + 1
#     return 0


# def px(hex_color: str) -> PatternFill:
#     return PatternFill("solid", fgColor=hex_color)


# def thin_border() -> Border:
#     s = Side(style="thin", color="D0D0D0")
#     return Border(left=s, right=s, top=s, bottom=s)


# def build_output_excel(mapped_df: pd.DataFrame) -> bytes:
#     wb = openpyxl.Workbook()
#     ws = wb.active
#     ws.title = "Diamond Data"

#     n          = len(mapped_df)
#     HDR        = 1
#     DS         = 2          # data start
#     DE         = DS + n - 1 # data end
#     SUM_ROW    = DE + 1
#     AVG_ROW    = DE + 2
#     LBL_ROW    = DE + 5     # FINAL / RATE / Difference

#     k_rng = f"$K${DS}:$K${DE}"
#     d_rng = f"$D${DS}:$D${DE}"
#     h_sum = f"$H${SUM_ROW}"

#     # ── Header row ──────────────────────────────────────────────────────────
#     for c, hdr in enumerate(OUTPUT_HEADERS, 1):
#         cell = ws.cell(row=HDR, column=c, value=hdr)
#         cell.fill      = px(CLR_HEADER)
#         cell.font      = Font(name="Arial", bold=True, color="FFFFFF", size=10)
#         cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
#         cell.border    = thin_border()
#     ws.row_dimensions[HDR].height = 36

#     # ── Data rows ────────────────────────────────────────────────────────────
#     for i, (_, row) in enumerate(mapped_df.iterrows()):
#         r = DS + i

#         # Raw value columns
#         raw = {
#             1: row.get("stock_no"),
#             2: row.get("certi_no"),
#             3: row.get("shape"),
#             4: row.get("carat"),
#             5: row.get("color"),
#             6: row.get("clarity"),
#             7: row.get("final"),
#             9: row.get("vdb"),
#             11: row.get("usa"),
#             13: row.get("with_tariff"),
#         }
#         # Formula columns
#         formulas = {
#             8:  f"=G{r}*D{r}",
#             10: f"=I{r}*D{r}",
#             12: f"=K{r}*D{r}",
#             14: f"=M{r}*D{r}",
#             15: f"=ROUND((K{r}*{h_sum})/SUMPRODUCT({k_rng},{d_rng}),2)",
#             16: f"=O{r}*D{r}",
#             17: f"=K{r}-O{r}",
#             18: f"=(Q{r}*100)/K{r}",
#             19: f"=O{r}*1.08",
#             20: f"=S{r}*D{r}",
#             21: f"=K{r}-S{r}",
#             22: f"=(U{r}*100)/O{r}",
#         }

#         for c in range(1, 23):
#             val = raw.get(c, formulas.get(c))
#             cell = ws.cell(row=r, column=c, value=val)
#             cell.border    = thin_border()
#             cell.alignment = Alignment(horizontal="center", vertical="center")
#             cell.font      = Font(name="Arial", size=10)
#             if c in COL_COLORS:
#                 cell.fill = px(COL_COLORS[c])
#             if c in (4, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22):
#                 cell.number_format = "0.00"

#         ws.row_dimensions[r].height = 18

#     # ── TOTAL row ─────────────────────────────────────────────────────────
#     total_fill = px("F2F2F2")
#     ws.cell(row=SUM_ROW, column=1, value="TOTAL").font = Font(name="Arial", bold=True, size=10)
#     for col, letter in {4:"D", 8:"H", 10:"J", 12:"L", 14:"N", 16:"P", 20:"T"}.items():
#         cell = ws.cell(row=SUM_ROW, column=col,
#                        value=f"=SUM({letter}{DS}:{letter}{DE})")
#         cell.font          = Font(name="Arial", bold=True, size=10)
#         cell.fill          = total_fill
#         cell.border        = thin_border()
#         cell.alignment     = Alignment(horizontal="center")
#         cell.number_format = "0.00"
#     ws.cell(row=SUM_ROW, column=1).fill   = total_fill
#     ws.cell(row=SUM_ROW, column=1).border = thin_border()
#     ws.cell(row=SUM_ROW, column=1).alignment = Alignment(horizontal="center")

#     # ── AVG row ───────────────────────────────────────────────────────────
#     avg_fill = px("E2EFDA")
#     ws.cell(row=AVG_ROW, column=1, value="AVG / CARAT").font = Font(name="Arial", bold=True, size=10)
#     for col, letter in {8:"H", 10:"J", 12:"L", 14:"N", 16:"P", 20:"T"}.items():
#         cell = ws.cell(row=AVG_ROW, column=col,
#                        value=f"={letter}{SUM_ROW}/D{SUM_ROW}")
#         cell.font          = Font(name="Arial", bold=True, size=10)
#         cell.fill          = avg_fill
#         cell.border        = thin_border()
#         cell.alignment     = Alignment(horizontal="center")
#         cell.number_format = "0.00"
#     ws.cell(row=AVG_ROW, column=1).fill   = avg_fill
#     ws.cell(row=AVG_ROW, column=1).border = thin_border()
#     ws.cell(row=AVG_ROW, column=1).alignment = Alignment(horizontal="center")

#     # ── Label block (FINAL / $ RATE / Difference) ─────────────────────────
#     lbl_data = [
#         (LBL_ROW,     "FINAL",      f"=H{AVG_ROW}"),
#         (LBL_ROW + 1, "$ RATE",     ""),
#         (LBL_ROW + 2, "Difference", f"=L{AVG_ROW}-N{AVG_ROW}"),
#     ]
#     for lrow, lbl, formula in lbl_data:
#         for c, val in ((7, lbl), (8, formula)):
#             cell = ws.cell(row=lrow, column=c, value=val)
#             cell.font      = Font(name="Arial", bold=True, size=10)
#             cell.border    = thin_border()
#             cell.alignment = Alignment(horizontal="center")
#             cell.number_format = "0.00"
#     # Highlight $ RATE cell for manual entry
#     rate_cell = ws.cell(row=LBL_ROW + 1, column=8)
#     rate_cell.fill = px("FFFF00")
#     rate_cell.font = Font(name="Arial", bold=True, color="FF0000", size=10)

#     # ── Column widths ──────────────────────────────────────────────────────
#     widths = [14, 14, 11, 8, 7, 9, 8, 11, 8, 11, 8, 11, 12, 12, 10, 12, 24, 22, 22, 22, 22, 22]
#     for i, w in enumerate(widths, 1):
#         ws.column_dimensions[get_column_letter(i)].width = w

#     ws.freeze_panes = "A2"

#     buf = io.BytesIO()
#     wb.save(buf)
#     buf.seek(0)
#     return buf.getvalue()


# # ─────────────────────────────────────────────────────────────────────────────
# #  HERO BANNER
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="hero">
#   <div class="hero-badge">✦ BRIGHT TRADING INDIA LLP</div>
#   <p class="hero-title">💎 DIAMOND COST OPTIMIZER</p>
#   <p class="hero-sub">Smart Cost Adjustment & Equal Profit Calculator</p>
# </div>
# """, unsafe_allow_html=True)

# # Color legend
# st.markdown("""
# <div class="legend-wrap">
#   <div class="pill"><div class="dot" style="background:#98F6EB"></div>Carat</div>
#   <div class="pill"><div class="dot" style="background:#FFCF37"></div>Final / Final Amt</div>
#   <div class="pill"><div class="dot" style="background:#5991D5"></div>VDB / VDB Amt</div>
#   <div class="pill"><div class="dot" style="background:#F79B4F"></div>USA / USA Amt</div>
#   <div class="pill"><div class="dot" style="background:#9E89B9"></div>With Tariff / With Tariff Amt</div>
#   <div class="pill"><div class="dot" style="background:#98B850"></div>Adj Cost Group</div>
#   <div class="pill"><div class="dot" style="background:#CB6967"></div>Tariff Group</div>
# </div>
# """, unsafe_allow_html=True)

# # ─────────────────────────────────────────────────────────────────────────────
# #  STEP 1 – UPLOAD
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <style>
# /* Pull the uploader widget up to sit flush inside the step-1 card */
# div[data-testid="stFileUploader"] {
#     background: #1a1d2e !important;
#     border: 2px dashed #3a3d5c !important;
#     border-radius: 12px !important;
#     padding: 1.2rem 1.4rem !important;
#     margin-top: 0 !important;
# }
# div[data-testid="stFileUploader"]:hover {
#     border-color: #6366f1 !important;
#     background: #1e2040 !important;
# }
# /* Label above the uploader */
# div[data-testid="stFileUploader"] label {
#     color: #94a3b8 !important;
#     font-size: .88rem !important;
# }
# /* "Upload" button inside uploader */
# div[data-testid="stFileUploader"] button {
#     background: linear-gradient(135deg,#6366f1,#8b5cf6) !important;
#     color: white !important;
#     border: none !important;
#     border-radius: 8px !important;
#     font-weight: 600 !important;
#     padding: .45rem 1.1rem !important;
# }
# div[data-testid="stFileUploader"] button:hover {
#     opacity: .88 !important;
# }
# /* Caption text (200MB per file…) */
# div[data-testid="stFileUploader"] small,
# div[data-testid="stFileUploader"] span {
#     color: #ffffff !important;
#     font-size: .78rem !important;
# }
# /* Collapse gap between card and uploader element */
# .upload-card-wrap + div { margin-top: -0.5rem !important; }
# </style>
# <div class="step-card upload-card-wrap">
#   <div class="step-header">
#     <span class="step-num">1</span>
#     <span class="step-title">Upload Party File</span>
#   </div>
#   <p style="color:#64748b;font-size:.87rem;margin:0 0 .9rem 42px;">
#     Accepts <b style="color:#a78bfa">CSV, XLSX, XLS, XLSM, TSV</b> — any column order, any layout.
#   </p>
# </div>
# """, unsafe_allow_html=True)

# party_file = st.file_uploader(
#     "Drop your file here or click Upload",
#     type=["csv", "xlsx", "xls", "xlsm", "tsv"],
#     key="party_file",
#     label_visibility="collapsed",
# )

# if party_file is None:
#     st.markdown('<div class="banner banner-info" style="margin-top:.6rem">👆 Upload the party file above to begin.</div>',
#                 unsafe_allow_html=True)
#     st.stop()

# # ── Read & normalise ──────────────────────────────────────────────────────────
# raw_df = read_uploaded_file(party_file)
# if raw_df is None:
#     st.stop()

# h_row = detect_header_row(raw_df)
# df = raw_df.copy()
# df.columns = raw_df.iloc[h_row]
# df = df.iloc[h_row + 1:].reset_index(drop=True)
# df.columns = [str(c).strip() if pd.notna(c) else f"Col_{i}" for i, c in enumerate(df.columns)]
# df = df.dropna(how="all")

# # Stats
# st.markdown(f"""
# <div class="stats-row">
#   <div class="stat-box"><div class="stat-val">{len(df)}</div><div class="stat-label">Rows Detected</div></div>
#   <div class="stat-box"><div class="stat-val">{len(df.columns)}</div><div class="stat-label">Columns Found</div></div>
#   <div class="stat-box"><div class="stat-val">{party_file.name.rsplit(".",1)[-1].upper()}</div><div class="stat-label">File Format</div></div>
# </div>
# """, unsafe_allow_html=True)

# with st.expander("👀 Preview raw data (first 5 rows)"):
#     st.dataframe(df.head(5), use_container_width=True)

# # ─────────────────────────────────────────────────────────────────────────────
# #  STEP 2 – COLUMN MAPPING
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="step-card">
#   <div class="step-header">
#     <span class="step-num">2</span>
#     <span class="step-title">Map Columns</span>
#   </div>
# </div>
# """, unsafe_allow_html=True)
# st.caption("Select which column in the party file corresponds to each required field. Auto-matched where possible.")

# available = ["— skip / not available —"] + list(df.columns)
# mapping   = {}

# cols_left, cols_right = st.columns(2)
# for idx, (key, label) in enumerate(REQUIRED_COLS):
#     default = smart_default(key, available)
#     container = cols_left if idx % 2 == 0 else cols_right
#     with container:
#         choice = st.selectbox(f"**{label}**", options=available,
#                               index=default, key=f"map_{key}")
#         mapping[key] = None if choice.startswith("—") else choice

# # Validate mandatory
# mandatory = ["carat", "final", "vdb", "usa", "with_tariff"]
# missing   = [next(l for kk, l in REQUIRED_COLS if kk == k)
#              for k in mandatory if not mapping.get(k)]
# if missing:
#     st.markdown(f'<div class="banner banner-warn">⚠️ Please map required columns: <b>{", ".join(missing)}</b></div>',
#                 unsafe_allow_html=True)
#     st.stop()

# # ─────────────────────────────────────────────────────────────────────────────
# #  STEP 3 – PREVIEW & DOWNLOAD
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="step-card">
#   <div class="step-header">
#     <span class="step-num">3</span>
#     <span class="step-title">Preview & Download</span>
#   </div>
# </div>
# """, unsafe_allow_html=True)

# # Build mapped dataframe
# records = []
# for _, row in df.iterrows():
#     rec = {key: (row[col] if col else None) for key, col in mapping.items()}
#     records.append(rec)

# mapped_df = pd.DataFrame(records)
# for col in ["carat", "final", "vdb", "usa", "with_tariff"]:
#     mapped_df[col] = pd.to_numeric(mapped_df[col], errors="coerce")

# mapped_df = mapped_df.dropna(subset=["carat", "final", "vdb", "usa", "with_tariff"], how="all")
# n_stones  = len(mapped_df)

# if n_stones == 0:
#     st.markdown('<div class="banner banner-warn">⚠️ No valid rows found. Check your column mappings.</div>',
#                 unsafe_allow_html=True)
#     st.stop()

# st.markdown(f'<div class="banner banner-success">✅ <b>{n_stones} stones</b> ready to export.</div>',
#             unsafe_allow_html=True)

# preview_rename = {
#     "stock_no":"Stock ID.", "certi_no":"Certi No.", "shape":"Shape",
#     "carat":"Carat", "color":"Color", "clarity":"Clarity",
#     "final":"Final", "vdb":"VDB", "usa":"USA", "with_tariff":"With Tariff",
# }
# st.dataframe(mapped_df.rename(columns=preview_rename).head(10), use_container_width=True)

# # ── Generate Excel ────────────────────────────────────────────────────────────
# with st.spinner("Building Excel with formulas and colours…"):
#     xlsx_bytes  = build_output_excel(mapped_df)
#     out_filename = safe_filename(party_file.name, n_stones)

# st.markdown('<div class="banner banner-success">💾 File ready — formulas & colours match your reference format exactly.</div>',
#             unsafe_allow_html=True)

# st.download_button(
#     label=f"⬇️  Download  {out_filename}",
#     data=xlsx_bytes,
#     file_name=out_filename,
#     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#     use_container_width=True,
# )

# # ── Formula reference ─────────────────────────────────────────────────────────
# with st.expander("📐 Formula & Colour Reference"):
#     st.markdown("""
# | Column | Colour | Source / Formula |
# |---|---|---|
# | Stock No., Certi No., Shape, Color, Clarity | — | From party file |
# | **Carat** | 🟦 Teal `#98F6EB` | From party file |
# | **Final**, **Final Amt** | 🟨 Gold `#FFCF37` | Final from party; Amt = `Final × Carat` |
# | **VDB**, **VDB Amt** | 🔵 Blue `#5991D5` | VDB from party; Amt = `VDB × Carat` |
# | **USA**, **USA Amt** | 🟠 Orange `#F79B4F` | USA from party; Amt = `USA × Carat` |
# | **With Tariff**, **Tariff Amt** | 🟣 Lavender `#9E89B9` | From party; Amt = `With Tariff × Carat` |
# | **Adj Cost** | 🟩 Green `#98B850` | `ROUND((USA × ΣFinalAmt) / SUMPRODUCT(USA,Carat), 2)` |
# | **Adj Amount** | 🟩 | `Adj Cost × Carat` |
# | **Profit $ (No Tariff)** | 🟩 | `USA − Adj Cost` |
# | **Profit % (No Tariff)** | 🟩 | `(Profit$ × 100) / USA` |
# | **Adj Cost With Tariff** | 🟥 Coral `#CB6967` | `Adj Cost × 1.08` |
# | **Adj Amt With Tariff** | 🟥 | `Adj Cost With Tariff × Carat` |
# | **Profit $ (With Tariff)** | 🟥 | `USA − Adj Cost With Tariff` |
# | **Profit % (With Tariff)** | 🟥 | `(Profit$ With Tariff × 100) / Adj Cost` |
# | **TOTAL row** | Light grey | SUM of each value column |
# | **AVG / CARAT row** | Light green | Each SUM ÷ Total Carat |
# | **Difference** | — | `USA Avg − Tariff Avg` |

# > 🟡 The **$ RATE** cell is highlighted yellow — fill it manually in Excel.
# """)


# import streamlit as st
# import pandas as pd
# import openpyxl
# from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
# from openpyxl.utils import get_column_letter
# import io
# import re
# import os

# # ─────────────────────────────────────────────────────────────────────────────
# #  PAGE CONFIG
# # ─────────────────────────────────────────────────────────────────────────────
# st.set_page_config(
#     page_title="💎 Diamond Formatter",
#     page_icon="💎",
#     layout="wide",
#     initial_sidebar_state="collapsed",
# )

# # ─────────────────────────────────────────────────────────────────────────────
# #  DARK THEME CSS
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <style>
# /* ── Base dark background ── */
# html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"] {
#     background-color: #0d0f14 !important;
#     color: #e2e8f0 !important;
# }
# [data-testid="stHeader"] { background: #0d0f14 !important; }
# [data-testid="stSidebar"] { background: #111318 !important; }
# section[data-testid="stMain"] > div { background: #0d0f14 !important; }

# /* ── Remove default padding top ── */
# .block-container { padding-top: 1.5rem !important; }

# /* ── Hero banner (animated) ── */
# .hero {
#     position: relative;
#     background: linear-gradient(135deg, #0f1629 0%, #1a0a2e 50%, #0a1628 100%);
#     background-size: 200% 200%;
#     animation: heroGradientShift 12s ease infinite;
#     border: 1px solid #2a2d3e;
#     border-radius: 20px;
#     padding: 2.4rem 2rem;
#     margin-bottom: 2rem;
#     overflow: hidden;
#     text-align: center;
#     box-shadow: 0 10px 40px rgba(0,0,0,.35);
# }
# @keyframes heroGradientShift {
#     0%   { background-position: 0% 50%; }
#     50%  { background-position: 100% 50%; }
#     100% { background-position: 0% 50%; }
# }
# /* floating glow orbs */
# .hero::before, .hero::after {
#     content: '';
#     position: absolute;
#     border-radius: 50%;
#     filter: blur(60px);
#     opacity: .55;
#     animation: heroFloat 8s ease-in-out infinite;
#     pointer-events: none;
# }
# .hero::before {
#     width: 240px; height: 240px;
#     background: rgba(99,102,241,.35);
#     top: -70px; left: 8%;
# }
# .hero::after {
#     width: 280px; height: 280px;
#     background: rgba(168,85,247,.28);
#     bottom: -90px; right: 8%;
#     animation-delay: -4s;
# }
# @keyframes heroFloat {
#     0%, 100% { transform: translateY(0) translateX(0); }
#     50%      { transform: translateY(-22px) translateX(16px); }
# }

# .hero-content { position: relative; z-index: 1; text-align: left; }

# /* Company name — shown FIRST, as a pulsing pill */
# .hero-company {
#     display: inline-block;
#     font-size: .92rem; font-weight: 700; letter-spacing: 2.6px;
#     color: #c4b5fd;
#     text-transform: uppercase;
#     padding: 7px 20px;
#     border: 1px solid rgba(167,139,250,.45);
#     border-radius: 30px;
#     background: rgba(99,102,241,.12);
#     margin: 0 0 1.1rem;
#     animation: heroFadeInDown .8s ease both,
#                heroPulseBorder 3s ease-in-out infinite 1s;
# }
# @keyframes heroFadeInDown {
#     from { opacity: 0; transform: translateY(-14px); }
#     to   { opacity: 1; transform: translateY(0); }
# }
# @keyframes heroPulseBorder {
#     0%, 100% { box-shadow: 0 0 0 rgba(167,139,250,0); }
#     50%      { box-shadow: 0 0 20px rgba(167,139,250,.4); }
# }

# /* Product / website name — shown SECOND, shimmering title */
# .hero-title {
#     display: flex;
#     align-items: center;
#     justify-content: flex-start;
#     gap: 12px;
#     margin: 0 0 .6rem;
#     animation: heroFadeInUp .9s ease .2s both;
# }
# .hero-gem-wrap {
#     position: relative;
#     display: inline-flex;
#     align-items: center;
#     justify-content: center;
# }
# .hero-title .gem {
#     display: inline-block;
#     font-size: 2rem;
#     line-height: 1;
#     transform-style: preserve-3d;
#     animation: heroGemSpin 4.5s ease-in-out infinite;
# }
# .hero-gem-wrap .spark {
#     position: absolute;
#     color: #7dd3fc;
#     font-size: .6rem;
#     opacity: 0;
# }
# .hero-gem-wrap .spark-1 { top: -6px;  right: -8px; animation: heroTwinkle 2.6s ease-in-out infinite; }
# .hero-gem-wrap .spark-2 { bottom: -4px; left: -10px; animation: heroTwinkle 2.6s ease-in-out infinite 1.1s; }
# @keyframes heroTwinkle {
#     0%, 100% { opacity: 0; transform: scale(.4) rotate(0deg); }
#     50%      { opacity: 1; transform: scale(1.1) rotate(25deg); }
# }
# @keyframes heroGemSpin {
#     0%   { transform: perspective(320px) rotateY(0deg) scale(1);
#            filter: drop-shadow(0 0 6px rgba(167,139,250,.45)); }
#     45%  { transform: perspective(320px) rotateY(180deg) scale(1.1);
#            filter: drop-shadow(0 0 16px rgba(56,189,248,.8)); }
#     55%  { transform: perspective(320px) rotateY(180deg) scale(1.1);
#            filter: drop-shadow(0 0 16px rgba(56,189,248,.8)); }
#     100% { transform: perspective(320px) rotateY(360deg) scale(1);
#            filter: drop-shadow(0 0 6px rgba(167,139,250,.45)); }
# }
# .hero-title .txt {
#     font-size: 2.3rem; font-weight: 800; letter-spacing: -.5px;
#     line-height: 1.15;
#     background: linear-gradient(90deg, #a78bfa, #6366f1, #38bdf8, #a78bfa);
#     background-size: 300% auto;
#     -webkit-background-clip: text; -webkit-text-fill-color: transparent;
#     background-clip: text;
#     animation: heroShimmer 6s linear infinite;
# }
# @keyframes heroShimmer {
#     to { background-position: 300% center; }
# }
# @keyframes heroFadeInUp {
#     from { opacity: 0; transform: translateY(18px); }
#     to   { opacity: 1; transform: translateY(0); }
# }

# .hero-sub {
#     color: #94a3b8; font-size: 1rem; margin: 0;
#     animation: heroFadeInUp .9s ease .45s both;
# }

# /* ── Color legend pills ── */
# .legend-wrap { display: flex; flex-wrap: wrap; gap: 8px; margin: 1rem 0 1.5rem; }
# .pill {
#     display: inline-flex; align-items: center; gap: 7px;
#     background: #1a1d27; border: 1px solid #2a2d3e;
#     border-radius: 20px; padding: 5px 12px; font-size: .78rem; color: #cbd5e1;
#     transition: transform .18s ease, border-color .18s ease;
# }
# .pill:hover {
#     transform: translateY(-2px);
#     border-color: #6366f1;
# }
# .dot { width: 12px; height: 12px; border-radius: 3px; flex-shrink: 0; }

# /* ── Step cards ── */
# .step-card {
#     background: #13161f;
#     border: 1px solid #1e2130;
#     border-radius: 12px;
#     padding: 1.5rem;
#     margin-bottom: 1rem;
# }
# .step-num {
#     display: inline-flex; align-items: center; justify-content: center;
#     width: 32px; height: 32px; border-radius: 50%;
#     background: linear-gradient(135deg, #6366f1, #8b5cf6);
#     color: white; font-weight: 700; font-size: .9rem;
#     margin-right: 10px; flex-shrink: 0;
# }
# .step-title { font-size: 1.1rem; font-weight: 600; color: #f1f5f9; }
# .step-header { display: flex; align-items: center; margin-bottom: 1rem; }

# /* ── Streamlit widget overrides (dark) ── */
# [data-testid="stFileUploader"] {
#     background: #13161f !important;
#     border: 2px dashed #2a2d3e !important;
#     border-radius: 10px !important;
# }
# [data-testid="stFileUploader"]:hover { border-color: #6366f1 !important; }
# .stSelectbox > div > div {
#     background: #1a1d27 !important;
#     border: 1px solid #2a2d3e !important;
#     color: #e2e8f0 !important;
#     border-radius: 8px !important;
# }
# .stSelectbox label { color: #cbd5e1 !important; font-size: .85rem !important; }
# [data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }

# /* ── Download button ── */
# .stDownloadButton button {
#     background: linear-gradient(135deg, #6366f1, #8b5cf6) !important;
#     color: white !important; border: none !important;
#     border-radius: 10px !important; font-weight: 600 !important;
#     padding: .75rem 1.5rem !important; font-size: 1rem !important;
#     transition: all .2s !important; box-shadow: 0 4px 20px rgba(99,102,241,.4) !important;
# }
# .stDownloadButton button:hover {
#     transform: translateY(-2px) !important;
#     box-shadow: 0 8px 30px rgba(99,102,241,.6) !important;
# }

# /* ── Success / info / warning banners ── */
# .banner {
#     border-radius: 10px; padding: .85rem 1.1rem;
#     font-size: .9rem; margin: .5rem 0; display: flex; align-items: center; gap: 10px;
# }
# .banner-success { background: rgba(34,197,94,.1); border: 1px solid rgba(34,197,94,.3); color: #86efac; }
# .banner-info    { background: rgba(99,102,241,.1); border: 1px solid rgba(99,102,241,.3); color: #a5b4fc; }
# .banner-warn    { background: rgba(234,179,8,.1);  border: 1px solid rgba(234,179,8,.3);  color: #fde68a; }

# /* ── Stats row ── */
# .stats-row { display: flex; gap: 12px; flex-wrap: wrap; margin: 1rem 0; }
# .stat-box {
#     background: #13161f; border: 1px solid #1e2130; border-radius: 10px;
#     padding: .75rem 1.25rem; flex: 1; min-width: 130px;
# }
# .stat-val { font-size: 1.6rem; font-weight: 700; color: #a78bfa; }
# .stat-label { font-size: .75rem; color: #64748b; text-transform: uppercase; letter-spacing: .5px; }

# /* ── Divider ── */
# hr { border-color: #1e2130 !important; margin: 1.5rem 0 !important; }

# /* ── Expander dark ── */
# [data-testid="stExpander"] {
#     background: #13161f !important; border: 1px solid #1e2130 !important;
#     border-radius: 10px !important;
# }
# summary { color: #94a3b8 !important; }

# /* ── Dataframe dark cells ── */
# .stDataFrame thead th { background: #1a1d27 !important; color: #a78bfa !important; }
# .stDataFrame tbody tr:nth-child(even) td { background: #13161f !important; }
# .stDataFrame tbody tr:nth-child(odd)  td { background: #0f111a !important; }
# </style>
# """, unsafe_allow_html=True)

# # ─────────────────────────────────────────────────────────────────────────────
# #  EXACT EXCEL COLORS (from Diamond_Formatted.xlsx)
# # ─────────────────────────────────────────────────────────────────────────────
# CLR_HEADER   = "1F4E79"   # deep navy  → col A-V header row
# CLR_CARAT    = "98F6EB"   # teal mint  → D (Carat)
# CLR_FINAL    = "FFCF37"   # golden     → G, H (Final, Final Amt)
# CLR_VDB      = "5991D5"   # steel blue → I, J (VDB, VDB Amt)
# CLR_USA      = "F79B4F"   # orange     → K, L (USA, USA Amt)
# CLR_TARIFF   = "9E89B9"   # lavender   → M, N (With Tariff, Tariff Amt)
# CLR_ADJ      = "98B850"   # sage green → O, P, Q, R (Adj Cost group)
# CLR_WTARIFF  = "CB6967"   # coral red  → S, T, U, V (With Tariff group)

# # col index → fill color (1-based)
# COL_COLORS = {
#     4:  CLR_CARAT,
#     7:  CLR_FINAL,   8:  CLR_FINAL,
#     9:  CLR_VDB,     10: CLR_VDB,
#     11: CLR_USA,     12: CLR_USA,
#     13: CLR_TARIFF,  14: CLR_TARIFF,
#     15: CLR_ADJ,     16: CLR_ADJ,  17: CLR_ADJ,  18: CLR_ADJ,
#     19: CLR_WTARIFF, 20: CLR_WTARIFF, 21: CLR_WTARIFF, 22: CLR_WTARIFF,
# }

# OUTPUT_HEADERS = [
#     "Stock ID.", "Certi No.", "Shape", "Carat", "Color", "Clarity",
#     "Final", "Final Amt.", "VDB", "VDB Amt.", "USA", "USA Amt.",
#     "With Tariff", "Tariff Amt.", "Adj Cost", "Adj Amount",
#     "Profit $ Without Tariff", "Profit % Without Tariff",
#     "Adj Cost With Tariff", "Adj Amt. With Tariff",
#     "Profit $ With Tariff", "Profit % With Tariff",
# ]

# REQUIRED_COLS = [
#     ("stock_no",    "Stock ID."),
#     ("certi_no",    "Certi No."),
#     ("shape",       "Shape"),
#     ("carat",       "Carat"),
#     ("color",       "Color"),
#     ("clarity",     "Clarity"),
#     ("final",       "Final"),
#     ("vdb",         "VDB"),
#     ("usa",         "USA"),
#     ("with_tariff", "With Tariff"),
# ]

# # ─────────────────────────────────────────────────────────────────────────────
# #  HELPERS
# # ─────────────────────────────────────────────────────────────────────────────
# def read_uploaded_file(f) -> pd.DataFrame | None:
#     name = f.name.lower()
#     try:
#         if name.endswith(".csv"):
#             return pd.read_csv(f)
#         elif name.endswith((".xlsx", ".xls", ".xlsm")):
#             return pd.read_excel(f, header=None)
#         elif name.endswith(".tsv"):
#             return pd.read_csv(f, sep="\t")
#         else:
#             st.error("Unsupported format. Please upload CSV, XLSX, XLS, XLSM, or TSV.")
#             return None
#     except Exception as e:
#         st.error(f"Could not read file: {e}")
#         return None


# def detect_header_row(df: pd.DataFrame) -> int:
#     for i, row in df.iterrows():
#         vals = [str(v).strip() for v in row if pd.notna(v) and str(v).strip()]
#         if len(vals) >= 4 and not all(
#             re.sub(r'[.\-]', '', v).isdigit() for v in vals
#         ):
#             return i
#     return 0


# def safe_filename(raw_name: str, n_stones: int) -> str:
#     """PartyFile_19_Stones.xlsx"""
#     stem = os.path.splitext(raw_name)[0]
#     stem = re.sub(r'[^\w\s\-]', '', stem).strip()
#     stem = re.sub(r'\s+', '_', stem)
#     return f"{stem}_{n_stones}_Stones.xlsx"


# def smart_default(key: str, cols: list) -> int:
#     hints = {
#         "stock_no":    ["stock", "pexkt", "packet", "pkt", "id", "lot"],
#         "certi_no":    ["certi", "cert", "gia", "igi", "report"],
#         "shape":       ["shape"],
#         "carat":       ["carat", "ct", "weight"],
#         "color":       ["color", "colour"],
#         "clarity":     ["clarity", "clar"],
#         "final":       ["final", "rap%", "disc%", "back%"],
#         "vdb":         ["vdb"],
#         "usa":         ["usa"],
#         "with_tariff": ["tariff", "tarrif", "with tariff", "wtariff"],
#     }
#     keywords = hints.get(key, [])
#     lower_cols = [c.lower() for c in cols[1:]]
#     for kw in keywords:
#         for i, lc in enumerate(lower_cols):
#             if kw in lc:
#                 return i + 1
#     return 0


# def px(hex_color: str) -> PatternFill:
#     return PatternFill("solid", fgColor=hex_color)


# def thin_border() -> Border:
#     s = Side(style="thin", color="D0D0D0")
#     return Border(left=s, right=s, top=s, bottom=s)


# def build_output_excel(mapped_df: pd.DataFrame) -> bytes:
#     wb = openpyxl.Workbook()
#     ws = wb.active
#     ws.title = "Diamond Data"

#     n          = len(mapped_df)
#     HDR        = 1
#     DS         = 2          # data start
#     DE         = DS + n - 1 # data end
#     SUM_ROW    = DE + 1
#     AVG_ROW    = DE + 2
#     LBL_ROW    = DE + 5     # FINAL / RATE / Difference

#     k_rng = f"$K${DS}:$K${DE}"
#     d_rng = f"$D${DS}:$D${DE}"
#     h_sum = f"$H${SUM_ROW}"

#     # ── Header row ──────────────────────────────────────────────────────────
#     for c, hdr in enumerate(OUTPUT_HEADERS, 1):
#         cell = ws.cell(row=HDR, column=c, value=hdr)
#         cell.fill      = px(CLR_HEADER)
#         cell.font      = Font(name="Arial", bold=True, color="FFFFFF", size=10)
#         cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
#         cell.border    = thin_border()
#     ws.row_dimensions[HDR].height = 36

#     # ── Data rows ────────────────────────────────────────────────────────────
#     for i, (_, row) in enumerate(mapped_df.iterrows()):
#         r = DS + i

#         # Raw value columns
#         raw = {
#             1: row.get("stock_no"),
#             2: row.get("certi_no"),
#             3: row.get("shape"),
#             4: row.get("carat"),
#             5: row.get("color"),
#             6: row.get("clarity"),
#             7: row.get("final"),
#             9: row.get("vdb"),
#             11: row.get("usa"),
#             13: row.get("with_tariff"),
#         }
#         # Formula columns
#         formulas = {
#             8:  f"=ROUND(G{r}*D{r},2)",
#             10: f"=ROUND(I{r}*D{r},2)",
#             12: f"=ROUND(K{r}*D{r},2)",
#             14: f"=ROUND(M{r}*D{r},2)",
#             15: f"=ROUND((K{r}*{h_sum})/SUMPRODUCT({k_rng},{d_rng}),2)",
#             16: f"=ROUND(O{r}*D{r},2)",
#             17: f"=ROUND(K{r}-O{r},2)",
#             18: f"=ROUND((Q{r}*100)/K{r},2)",
#             19: f"=ROUND(O{r}*1.08,2)",
#             20: f"=ROUND(S{r}*D{r},2)",
#             21: f"=ROUND(K{r}-S{r},2)",
#             22: f"=ROUND((U{r}*100)/O{r},2)",
#         }

#         for c in range(1, 23):
#             val = raw.get(c, formulas.get(c))
#             cell = ws.cell(row=r, column=c, value=val)
#             cell.border    = thin_border()
#             cell.alignment = Alignment(horizontal="center", vertical="center")
#             cell.font      = Font(name="Arial", size=10)
#             if c in COL_COLORS:
#                 cell.fill = px(COL_COLORS[c])
#             if c in (4, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22):
#                 cell.number_format = "0.00"

#         ws.row_dimensions[r].height = 18

#     # ── TOTAL row ─────────────────────────────────────────────────────────
#     total_fill = px("F2F2F2")
#     ws.cell(row=SUM_ROW, column=1, value="TOTAL").font = Font(name="Arial", bold=True, size=10)
#     for col, letter in {4:"D", 8:"H", 10:"J", 12:"L", 14:"N", 16:"P", 20:"T"}.items():
#         cell = ws.cell(row=SUM_ROW, column=col,
#                        value=f"=SUM({letter}{DS}:{letter}{DE})")
#         cell.font          = Font(name="Arial", bold=True, size=10)
#         cell.fill          = total_fill
#         cell.border        = thin_border()
#         cell.alignment     = Alignment(horizontal="center")
#         cell.number_format = "0.00"
#     ws.cell(row=SUM_ROW, column=1).fill   = total_fill
#     ws.cell(row=SUM_ROW, column=1).border = thin_border()
#     ws.cell(row=SUM_ROW, column=1).alignment = Alignment(horizontal="center")

#     # ── AVG row ───────────────────────────────────────────────────────────
#     avg_fill = px("E2EFDA")
#     ws.cell(row=AVG_ROW, column=1, value="AVG / CARAT").font = Font(name="Arial", bold=True, size=10)
#     for col, letter in {8:"H", 10:"J", 12:"L", 14:"N", 16:"P", 20:"T"}.items():
#         cell = ws.cell(row=AVG_ROW, column=col,
#                        value=f"=ROUND({letter}{SUM_ROW}/D{SUM_ROW},2)")
#         cell.font          = Font(name="Arial", bold=True, size=10)
#         cell.fill          = avg_fill
#         cell.border        = thin_border()
#         cell.alignment     = Alignment(horizontal="center")
#         cell.number_format = "0.00"
#     ws.cell(row=AVG_ROW, column=1).fill   = avg_fill
#     ws.cell(row=AVG_ROW, column=1).border = thin_border()
#     ws.cell(row=AVG_ROW, column=1).alignment = Alignment(horizontal="center")

#     # ── Label block (FINAL / $ RATE / Difference) ─────────────────────────
#     lbl_data = [
#         (LBL_ROW,     "FINAL",      f"=ROUND(H{AVG_ROW},2)"),
#         (LBL_ROW + 1, "$ RATE",     ""),
#         (LBL_ROW + 2, "Difference", f"=L{AVG_ROW}-N{AVG_ROW}"),
#     ]
#     for lrow, lbl, formula in lbl_data:
#         for c, val in ((7, lbl), (8, formula)):
#             cell = ws.cell(row=lrow, column=c, value=val)
#             cell.font      = Font(name="Arial", bold=True, size=10)
#             cell.border    = thin_border()
#             cell.alignment = Alignment(horizontal="center")
#             cell.number_format = "0.00"
#     # Highlight $ RATE cell for manual entry
#     rate_cell = ws.cell(row=LBL_ROW + 1, column=8)
#     rate_cell.fill = px("FFFF00")
#     rate_cell.font = Font(name="Arial", bold=True, color="FF0000", size=10)

#     # ── Column widths ──────────────────────────────────────────────────────
#     widths = [14, 14, 11, 8, 7, 9, 8, 11, 8, 11, 8, 11, 12, 12, 10, 12, 24, 22, 22, 22, 22, 22]
#     for i, w in enumerate(widths, 1):
#         ws.column_dimensions[get_column_letter(i)].width = w

#     ws.freeze_panes = "A2"

#     buf = io.BytesIO()
#     wb.save(buf)
#     buf.seek(0)
#     return buf.getvalue()


# # ─────────────────────────────────────────────────────────────────────────────
# #  HERO BANNER  (Company name first, product/website name second, animated)
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="hero">
#   <div class="hero-content">
#     <p class="hero-company">✦ BRIGHT TRADING INDIA LLP</p>
#     <p class="hero-title">
#       <span class="hero-gem-wrap">
#         <span class="gem">💎</span>
#         <span class="spark spark-1">✦</span>
#         <span class="spark spark-2">✦</span>
#       </span>
#       <span class="txt">Diamond Cost Optimizer</span>
#     </p>
#     <p class="hero-sub">Smart Cost Adjustment &amp; Equal Profit Calculator</p>
#   </div>
# </div>
# """, unsafe_allow_html=True)

# # Color legend
# st.markdown("""
# <div class="legend-wrap">
#   <div class="pill"><div class="dot" style="background:#98F6EB"></div>Carat</div>
#   <div class="pill"><div class="dot" style="background:#FFCF37"></div>Final / Final Amt</div>
#   <div class="pill"><div class="dot" style="background:#5991D5"></div>VDB / VDB Amt</div>
#   <div class="pill"><div class="dot" style="background:#F79B4F"></div>USA / USA Amt</div>
#   <div class="pill"><div class="dot" style="background:#9E89B9"></div>With Tariff / With Tariff Amt</div>
#   <div class="pill"><div class="dot" style="background:#98B850"></div>Adj Cost Group</div>
#   <div class="pill"><div class="dot" style="background:#CB6967"></div>Tariff Group</div>
# </div>
# """, unsafe_allow_html=True)

# # ─────────────────────────────────────────────────────────────────────────────
# #  STEP 1 – UPLOAD
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <style>
# /* Pull the uploader widget up to sit flush inside the step-1 card */
# div[data-testid="stFileUploader"] {
#     background: #1a1d2e !important;
#     border: 2px dashed #3a3d5c !important;
#     border-radius: 12px !important;
#     padding: 1.2rem 1.4rem !important;
#     margin-top: 0 !important;
# }
# div[data-testid="stFileUploader"]:hover {
#     border-color: #6366f1 !important;
#     background: #1e2040 !important;
# }
# /* Label above the uploader */
# div[data-testid="stFileUploader"] label {
#     color: #94a3b8 !important;
#     font-size: .88rem !important;
# }
# /* "Upload" button inside uploader */
# div[data-testid="stFileUploader"] button {
#     background: linear-gradient(135deg,#6366f1,#8b5cf6) !important;
#     color: white !important;
#     border: none !important;
#     border-radius: 8px !important;
#     font-weight: 600 !important;
#     padding: .45rem 1.1rem !important;
# }
# div[data-testid="stFileUploader"] button:hover {
#     opacity: .88 !important;
# }
# /* Caption text (200MB per file…) */
# div[data-testid="stFileUploader"] small,
# div[data-testid="stFileUploader"] span {
#     color: #ffffff !important;
#     font-size: .78rem !important;
# }
# /* Collapse gap between card and uploader element */
# .upload-card-wrap + div { margin-top: -0.5rem !important; }
# </style>
# <div class="step-card upload-card-wrap">
#   <div class="step-header">
#     <span class="step-num">1</span>
#     <span class="step-title">Upload Party File</span>
#   </div>
#   <p style="color:#64748b;font-size:.87rem;margin:0 0 .9rem 42px;">
#     Accepts <b style="color:#a78bfa">CSV, XLSX, XLS, XLSM, TSV</b> — any column order, any layout.
#   </p>
# </div>
# """, unsafe_allow_html=True)

# party_file = st.file_uploader(
#     "Drop your file here or click Upload",
#     type=["csv", "xlsx", "xls", "xlsm", "tsv"],
#     key="party_file",
#     label_visibility="collapsed",
# )

# if party_file is None:
#     st.markdown('<div class="banner banner-info" style="margin-top:.6rem">👆 Upload the party file above to begin.</div>',
#                 unsafe_allow_html=True)
#     st.stop()

# # ── Read & normalise ──────────────────────────────────────────────────────────
# raw_df = read_uploaded_file(party_file)
# if raw_df is None:
#     st.stop()

# h_row = detect_header_row(raw_df)
# df = raw_df.copy()
# df.columns = raw_df.iloc[h_row]
# df = df.iloc[h_row + 1:].reset_index(drop=True)
# df.columns = [str(c).strip() if pd.notna(c) else f"Col_{i}" for i, c in enumerate(df.columns)]
# df = df.dropna(how="all")

# # Stats
# st.markdown(f"""
# <div class="stats-row">
#   <div class="stat-box"><div class="stat-val">{len(df)}</div><div class="stat-label">Rows Detected</div></div>
#   <div class="stat-box"><div class="stat-val">{len(df.columns)}</div><div class="stat-label">Columns Found</div></div>
#   <div class="stat-box"><div class="stat-val">{party_file.name.rsplit(".",1)[-1].upper()}</div><div class="stat-label">File Format</div></div>
# </div>
# """, unsafe_allow_html=True)

# with st.expander("👀 Preview raw data (first 5 rows)"):
#     st.dataframe(df.head(5), use_container_width=True)

# # ─────────────────────────────────────────────────────────────────────────────
# #  STEP 2 – COLUMN MAPPING
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="step-card">
#   <div class="step-header">
#     <span class="step-num">2</span>
#     <span class="step-title">Map Columns</span>
#   </div>
# </div>
# """, unsafe_allow_html=True)
# st.caption("Select which column in the party file corresponds to each required field. Auto-matched where possible.")

# available = ["— skip / not available —"] + list(df.columns)
# mapping   = {}

# cols_left, cols_right = st.columns(2)
# for idx, (key, label) in enumerate(REQUIRED_COLS):
#     default = smart_default(key, available)
#     container = cols_left if idx % 2 == 0 else cols_right
#     with container:
#         choice = st.selectbox(f"**{label}**", options=available,
#                               index=default, key=f"map_{key}")
#         mapping[key] = None if choice.startswith("—") else choice

# # Validate mandatory
# mandatory = ["carat", "final", "vdb", "usa", "with_tariff"]
# missing   = [next(l for kk, l in REQUIRED_COLS if kk == k)
#              for k in mandatory if not mapping.get(k)]
# if missing:
#     st.markdown(f'<div class="banner banner-warn">⚠️ Please map required columns: <b>{", ".join(missing)}</b></div>',
#                 unsafe_allow_html=True)
#     st.stop()

# # ─────────────────────────────────────────────────────────────────────────────
# #  STEP 3 – PREVIEW & DOWNLOAD
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="step-card">
#   <div class="step-header">
#     <span class="step-num">3</span>
#     <span class="step-title">Preview & Download</span>
#   </div>
# </div>
# """, unsafe_allow_html=True)

# # Build mapped dataframe
# records = []
# for _, row in df.iterrows():
#     rec = {key: (row[col] if col else None) for key, col in mapping.items()}
#     records.append(rec)

# mapped_df = pd.DataFrame(records)
# for col in ["carat", "final", "vdb", "usa", "with_tariff"]:
#     mapped_df[col] = pd.to_numeric(mapped_df[col], errors="coerce")

# mapped_df = mapped_df.dropna(subset=["carat", "final", "vdb", "usa", "with_tariff"], how="all")
# n_stones  = len(mapped_df)

# if n_stones == 0:
#     st.markdown('<div class="banner banner-warn">⚠️ No valid rows found. Check your column mappings.</div>',
#                 unsafe_allow_html=True)
#     st.stop()

# st.markdown(f'<div class="banner banner-success">✅ <b>{n_stones} stones</b> ready to export.</div>',
#             unsafe_allow_html=True)

# preview_rename = {
#     "stock_no":"Stock ID.", "certi_no":"Certi No.", "shape":"Shape",
#     "carat":"Carat", "color":"Color", "clarity":"Clarity",
#     "final":"Final", "vdb":"VDB", "usa":"USA", "with_tariff":"With Tariff",
# }
# st.dataframe(mapped_df.rename(columns=preview_rename).head(10), use_container_width=True)

# # ── Generate Excel ────────────────────────────────────────────────────────────
# with st.spinner("Building Excel with formulas and colours…"):
#     xlsx_bytes  = build_output_excel(mapped_df)
#     out_filename = safe_filename(party_file.name, n_stones)

# st.markdown('<div class="banner banner-success">💾 File ready — formulas & colours match your reference format exactly.</div>',
#             unsafe_allow_html=True)

# st.download_button(
#     label=f"⬇️  Download  {out_filename}",
#     data=xlsx_bytes,
#     file_name=out_filename,
#     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#     use_container_width=True,
# )

# # ── Formula reference ─────────────────────────────────────────────────────────
# with st.expander("📐 Formula & Colour Reference"):
#     st.markdown("""
# | Column | Colour | Source / Formula |
# |---|---|---|
# | Stock No., Certi No., Shape, Color, Clarity | — | From party file |
# | **Carat** | 🟦 Teal `#98F6EB` | From party file |
# | **Final**, **Final Amt** | 🟨 Gold `#FFCF37` | Final from party; Amt = `Final × Carat` |
# | **VDB**, **VDB Amt** | 🔵 Blue `#5991D5` | VDB from party; Amt = `VDB × Carat` |
# | **USA**, **USA Amt** | 🟠 Orange `#F79B4F` | USA from party; Amt = `USA × Carat` |
# | **With Tariff**, **Tariff Amt** | 🟣 Lavender `#9E89B9` | From party; Amt = `With Tariff × Carat` |
# | **Adj Cost** | 🟩 Green `#98B850` | `ROUND((USA × ΣFinalAmt) / SUMPRODUCT(USA,Carat), 2)` |
# | **Adj Amount** | 🟩 | `Adj Cost × Carat` |
# | **Profit $ (No Tariff)** | 🟩 | `USA − Adj Cost` |
# | **Profit % (No Tariff)** | 🟩 | `(Profit$ × 100) / USA` |
# | **Adj Cost With Tariff** | 🟥 Coral `#CB6967` | `Adj Cost × 1.08` |
# | **Adj Amt With Tariff** | 🟥 | `Adj Cost With Tariff × Carat` |
# | **Profit $ (With Tariff)** | 🟥 | `USA − Adj Cost With Tariff` |
# | **Profit % (With Tariff)** | 🟥 | `(Profit$ With Tariff × 100) / Adj Cost` |
# | **TOTAL row** | Light grey | SUM of each value column |
# | **AVG / CARAT row** | Light green | Each SUM ÷ Total Carat |
# | **Difference** | — | `USA Avg − Tariff Avg` |

# > 🟡 The **$ RATE** cell is highlighted yellow — fill it manually in Excel.
# """)


# import streamlit as st
# import pandas as pd
# import openpyxl
# from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
# from openpyxl.utils import get_column_letter
# import io
# import re
# import os

# # ─────────────────────────────────────────────────────────────────────────────
# #  PAGE CONFIG
# # ─────────────────────────────────────────────────────────────────────────────
# st.set_page_config(
#     page_title="💎 Diamond Formatter",
#     page_icon="💎",
#     layout="wide",
#     initial_sidebar_state="collapsed",
# )

# # ─────────────────────────────────────────────────────────────────────────────
# #  DARK THEME CSS
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <style>
# html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"] {
#     background-color: #0d0f14 !important; color: #e2e8f0 !important;
# }
# [data-testid="stHeader"] { background: #0d0f14 !important; }
# [data-testid="stSidebar"] { background: #111318 !important; }
# section[data-testid="stMain"] > div { background: #0d0f14 !important; }
# .block-container { padding-top: 1.5rem !important; }
# .hero {
#     position: relative;
#     background: linear-gradient(135deg, #0f1629 0%, #1a0a2e 50%, #0a1628 100%);
#     background-size: 200% 200%;
#     animation: heroGradientShift 12s ease infinite;
#     border: 1px solid #2a2d3e; border-radius: 20px;
#     padding: 2.4rem 2rem; margin-bottom: 2rem; overflow: hidden;
#     text-align: center; box-shadow: 0 10px 40px rgba(0,0,0,.35);
# }
# @keyframes heroGradientShift {
#     0%   { background-position: 0% 50%; }
#     50%  { background-position: 100% 50%; }
#     100% { background-position: 0% 50%; }
# }
# .hero::before, .hero::after {
#     content: ''; position: absolute; border-radius: 50%;
#     filter: blur(60px); opacity: .55;
#     animation: heroFloat 8s ease-in-out infinite; pointer-events: none;
# }
# .hero::before { width:240px;height:240px;background:rgba(99,102,241,.35);top:-70px;left:8%; }
# .hero::after  { width:280px;height:280px;background:rgba(168,85,247,.28);bottom:-90px;right:8%;animation-delay:-4s; }
# @keyframes heroFloat {
#     0%,100% { transform:translateY(0) translateX(0); }
#     50%     { transform:translateY(-22px) translateX(16px); }
# }
# .hero-content { position:relative;z-index:1;text-align:left; }
# .hero-company {
#     display:inline-block;font-size:.92rem;font-weight:700;letter-spacing:2.6px;
#     color:#c4b5fd;text-transform:uppercase;padding:7px 20px;
#     border:1px solid rgba(167,139,250,.45);border-radius:30px;
#     background:rgba(99,102,241,.12);margin:0 0 1.1rem;
#     animation:heroFadeInDown .8s ease both,heroPulseBorder 3s ease-in-out infinite 1s;
# }
# @keyframes heroFadeInDown { from{opacity:0;transform:translateY(-14px)} to{opacity:1;transform:translateY(0)} }
# @keyframes heroPulseBorder {
#     0%,100% { box-shadow:0 0 0 rgba(167,139,250,0); }
#     50%     { box-shadow:0 0 20px rgba(167,139,250,.4); }
# }
# .hero-title { display:flex;align-items:center;justify-content:flex-start;gap:12px;margin:0 0 .6rem;animation:heroFadeInUp .9s ease .2s both; }
# .hero-gem-wrap { position:relative;display:inline-flex;align-items:center;justify-content:center; }
# .hero-title .gem { display:inline-block;font-size:2rem;line-height:1;transform-style:preserve-3d;animation:heroGemSpin 4.5s ease-in-out infinite; }
# .hero-gem-wrap .spark { position:absolute;color:#7dd3fc;font-size:.6rem;opacity:0; }
# .hero-gem-wrap .spark-1 { top:-6px;right:-8px;animation:heroTwinkle 2.6s ease-in-out infinite; }
# .hero-gem-wrap .spark-2 { bottom:-4px;left:-10px;animation:heroTwinkle 2.6s ease-in-out infinite 1.1s; }
# @keyframes heroTwinkle { 0%,100%{opacity:0;transform:scale(.4) rotate(0deg)} 50%{opacity:1;transform:scale(1.1) rotate(25deg)} }
# @keyframes heroGemSpin {
#     0%   { transform:perspective(320px) rotateY(0deg) scale(1);filter:drop-shadow(0 0 6px rgba(167,139,250,.45)); }
#     45%  { transform:perspective(320px) rotateY(180deg) scale(1.1);filter:drop-shadow(0 0 16px rgba(56,189,248,.8)); }
#     55%  { transform:perspective(320px) rotateY(180deg) scale(1.1);filter:drop-shadow(0 0 16px rgba(56,189,248,.8)); }
#     100% { transform:perspective(320px) rotateY(360deg) scale(1);filter:drop-shadow(0 0 6px rgba(167,139,250,.45)); }
# }
# .hero-title .txt {
#     font-size:2.3rem;font-weight:800;letter-spacing:-.5px;line-height:1.15;
#     background:linear-gradient(90deg,#a78bfa,#6366f1,#38bdf8,#a78bfa);
#     background-size:300% auto;
#     -webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;
#     animation:heroShimmer 6s linear infinite;
# }
# @keyframes heroShimmer { to{background-position:300% center} }
# @keyframes heroFadeInUp { from{opacity:0;transform:translateY(18px)} to{opacity:1;transform:translateY(0)} }
# .hero-sub { color:#94a3b8;font-size:1rem;margin:0;animation:heroFadeInUp .9s ease .45s both; }
# .legend-wrap { display:flex;flex-wrap:wrap;gap:8px;margin:1rem 0 1.5rem; }
# .pill {
#     display:inline-flex;align-items:center;gap:7px;
#     background:#1a1d27;border:1px solid #2a2d3e;
#     border-radius:20px;padding:5px 12px;font-size:.78rem;color:#cbd5e1;
#     transition:transform .18s ease,border-color .18s ease;
# }
# .pill:hover { transform:translateY(-2px);border-color:#6366f1; }
# .dot { width:12px;height:12px;border-radius:3px;flex-shrink:0; }
# .step-card { background:#13161f;border:1px solid #1e2130;border-radius:12px;padding:1.5rem;margin-bottom:1rem; }
# .step-num {
#     display:inline-flex;align-items:center;justify-content:center;
#     width:32px;height:32px;border-radius:50%;
#     background:linear-gradient(135deg,#6366f1,#8b5cf6);
#     color:white;font-weight:700;font-size:.9rem;margin-right:10px;flex-shrink:0;
# }
# .step-title { font-size:1.1rem;font-weight:600;color:#f1f5f9; }
# .step-header { display:flex;align-items:center;margin-bottom:1rem; }
# [data-testid="stFileUploader"] { background:#13161f !important;border:2px dashed #2a2d3e !important;border-radius:10px !important; }
# [data-testid="stFileUploader"]:hover { border-color:#6366f1 !important; }
# .stSelectbox > div > div { background:#1a1d27 !important;border:1px solid #2a2d3e !important;color:#e2e8f0 !important;border-radius:8px !important; }
# .stSelectbox label { color:#cbd5e1 !important;font-size:.85rem !important; }
# .stDownloadButton button {
#     background:linear-gradient(135deg,#6366f1,#8b5cf6) !important;
#     color:white !important;border:none !important;border-radius:10px !important;
#     font-weight:600 !important;padding:.75rem 1.5rem !important;
#     font-size:1rem !important;transition:all .2s !important;
#     box-shadow:0 4px 20px rgba(99,102,241,.4) !important;
# }
# .stDownloadButton button:hover { transform:translateY(-2px) !important;box-shadow:0 8px 30px rgba(99,102,241,.6) !important; }
# .banner { border-radius:10px;padding:.85rem 1.1rem;font-size:.9rem;margin:.5rem 0;display:flex;align-items:center;gap:10px; }
# .banner-success { background:rgba(34,197,94,.1);border:1px solid rgba(34,197,94,.3);color:#86efac; }
# .banner-info    { background:rgba(99,102,241,.1);border:1px solid rgba(99,102,241,.3);color:#a5b4fc; }
# .banner-warn    { background:rgba(234,179,8,.1); border:1px solid rgba(234,179,8,.3); color:#fde68a; }
# .stats-row { display:flex;gap:12px;flex-wrap:wrap;margin:1rem 0; }
# .stat-box { background:#13161f;border:1px solid #1e2130;border-radius:10px;padding:.75rem 1.25rem;flex:1;min-width:130px; }
# .stat-val { font-size:1.6rem;font-weight:700;color:#a78bfa; }
# .stat-label { font-size:.75rem;color:#64748b;text-transform:uppercase;letter-spacing:.5px; }
# hr { border-color:#1e2130 !important;margin:1.5rem 0 !important; }
# [data-testid="stExpander"] { background:#13161f !important;border:1px solid #1e2130 !important;border-radius:10px !important; }
# summary { color:#94a3b8 !important; }
# </style>
# """, unsafe_allow_html=True)

# # ─────────────────────────────────────────────────────────────────────────────
# #  COLOR CONSTANTS
# #  Matches Image 2 exactly:
# #  - Cols A-O : existing group colors (navy header, teal, gold, blue, orange, lavender)
# #  - Cols P-Y : YELLOW (FFFF00) — boss requirement
# # ─────────────────────────────────────────────────────────────────────────────
# CLR_HEADER  = "1F4E79"   # deep navy header
# CLR_CARAT   = "98F6EB"   # teal   → col D (4)
# CLR_FINAL   = "FFCF37"   # gold   → H,I (8,9)
# CLR_VDB     = "5991D5"   # blue   → J,K (10,11)
# CLR_USA     = "F79B4F"   # orange → L,M (12,13)
# CLR_TARIFF  = "9E89B9"   # lavender → N,O (14,15)
# CLR_YELLOW  = "FFFF00"   # yellow → P-Y (16-25)

# # col index (1-based) → fill color for DATA cells
# COL_COLORS = {
#     4:  CLR_CARAT,
#     8:  CLR_FINAL,   9:  CLR_FINAL,
#     10: CLR_VDB,     11: CLR_VDB,
#     12: CLR_USA,     13: CLR_USA,
#     14: CLR_TARIFF,  15: CLR_TARIFF,
# }
# # cols 16-26 all yellow
# for _c in range(16, 27):
#     COL_COLORS[_c] = CLR_YELLOW

# # ─────────────────────────────────────────────────────────────────────────────
# #  OUTPUT HEADERS  (25 columns, matching Image 2)
# # ─────────────────────────────────────────────────────────────────────────────
# OUTPUT_HEADERS = [
#     "Stock ID.",        # A  1
#     "Certi No.",        # B  2
#     "Shape",            # C  3
#     "Carat",            # D  4
#     "Lab",              # E  5
#     "Color",            # F  6
#     "Clarity",          # G  7
#     "Final",            # H  8
#     "Amt",              # I  9
#     "VDB",              # J  10
#     "Amt",              # K  11
#     "USA",              # L  12
#     "Amt",              # M  13
#     "With Tariff",      # N  14
#     "Amt",              # O  15
#     "",                 # P  16  (yellow — unlabeled in source)
#     "",                 # Q  17
#     "",                 # R  18
#     "",                 # S  19
#     "",                 # T  20
#     "",                 # U  21
#     "",                 # V  22
#     "",                 # W  23
#     "",                 # X  24
#     "",                 # Y  25
#     "Total INR Amt",   # Z  26
# ]

# REQUIRED_COLS = [
#     ("stock_no",    "Stock ID."),
#     ("certi_no",    "Certi No."),
#     ("shape",       "Shape"),
#     ("carat",       "Carat"),
#     ("lab",         "Lab"),
#     ("color",       "Color"),
#     ("clarity",     "Clarity"),
#     ("final",       "Final"),
#     ("vdb",         "VDB"),
#     ("usa",         "USA"),
#     ("with_tariff", "With Tariff"),
# ]

# # ─────────────────────────────────────────────────────────────────────────────
# #  HELPERS
# # ─────────────────────────────────────────────────────────────────────────────
# def read_uploaded_file(f) -> pd.DataFrame | None:
#     name = f.name.lower()
#     try:
#         if name.endswith(".csv"):
#             return pd.read_csv(f)
#         elif name.endswith((".xlsx", ".xls", ".xlsm")):
#             return pd.read_excel(f, header=None)
#         elif name.endswith(".tsv"):
#             return pd.read_csv(f, sep="\t")
#         else:
#             st.error("Unsupported format. Please upload CSV, XLSX, XLS, XLSM, or TSV.")
#             return None
#     except Exception as e:
#         st.error(f"Could not read file: {e}")
#         return None


# def detect_header_row(df: pd.DataFrame) -> int:
#     for i, row in df.iterrows():
#         vals = [str(v).strip() for v in row if pd.notna(v) and str(v).strip()]
#         if len(vals) >= 4 and not all(
#             re.sub(r'[.\-]', '', v).isdigit() for v in vals
#         ):
#             return i
#     return 0


# def trim_data_rows(df: pd.DataFrame, stock_col: str | None, carat_col: str | None) -> pd.DataFrame:
#     """
#     Drop any trailing non-data rows (SUM totals, labels, blanks) that appear
#     after the real diamond records.

#     A row is treated as a SUM / footer row — and all rows from it onward are
#     dropped — when ANY of these conditions holds:

#     1. Carat column is mapped AND the carat cell has a numeric value BUT the
#        stock-ID cell is empty or NaN.  This is the classic pattern:
#            Stock  | Carat
#            (blank)| 28.41   ← SUM row
#     2. Carat column is mapped AND the carat value is ≥ 3× the median carat of
#        the rows seen so far (catches a huge SUM value even if stock is filled).
#     3. Stock-ID column is mapped AND the stock-ID cell is a purely numeric string
#        or a known summary keyword (TOTAL, AVG, SUM, RATE, FINAL, etc.).
#     """
#     if df.empty:
#         return df

#     summary_keywords = {"total", "avg", "average", "sum", "rate", "final",
#                         "difference", "subtotal", "done", "fianl"}

#     # Compute median carat from first 5 valid rows for outlier detection
#     median_carat = None
#     if carat_col and carat_col in df.columns:
#         sample_carats = pd.to_numeric(df[carat_col].head(20), errors="coerce").dropna()
#         if not sample_carats.empty:
#             median_carat = sample_carats.median()

#     cut_at = len(df)  # default: keep everything

#     for i, (_, row) in enumerate(df.iterrows()):
#         # ── condition 1: empty stock + non-empty carat ──
#         if stock_col and stock_col in df.columns and carat_col and carat_col in df.columns:
#             stock_val = row.get(stock_col)
#             carat_val = row.get(carat_col)
#             stock_empty = pd.isna(stock_val) or str(stock_val).strip() == ""
#             carat_numeric = pd.notna(carat_val) and str(carat_val).strip() != ""
#             if stock_empty and carat_numeric:
#                 cut_at = i
#                 break

#         # ── condition 2: carat value is a huge outlier (SUM) ──
#         if carat_col and carat_col in df.columns and median_carat:
#             carat_val = pd.to_numeric(row.get(carat_col), errors="coerce")
#             if pd.notna(carat_val) and carat_val >= median_carat * 3:
#                 cut_at = i
#                 break

#         # ── condition 3: stock ID is a summary keyword ──
#         if stock_col and stock_col in df.columns:
#             stock_val = str(row.get(stock_col, "")).strip().lower()
#             if stock_val in summary_keywords:
#                 cut_at = i
#                 break

#     return df.iloc[:cut_at].reset_index(drop=True)


# def safe_filename(raw_name: str, n_stones: int) -> str:
#     stem = os.path.splitext(raw_name)[0]
#     stem = re.sub(r'[^\w\s\-]', '', stem).strip()
#     stem = re.sub(r'\s+', '_', stem)
#     return f"{stem}_{n_stones}_Stones.xlsx"


# def smart_default(key: str, cols: list) -> int:
#     hints = {
#         "stock_no":    ["stock", "pexkt", "packet", "pkt", "id", "lot"],
#         "certi_no":    ["certi", "cert", "gia", "igi", "report"],
#         "shape":       ["shape"],
#         "carat":       ["carat", "ct", "weight"],
#         "lab":         ["lab", "igi", "gia"],
#         "color":       ["color", "colour"],
#         "clarity":     ["clarity", "clar"],
#         "final":       ["final", "rap%", "disc%", "back%"],
#         "vdb":         ["vdb"],
#         "usa":         ["usa"],
#         "with_tariff": ["tariff", "tarrif", "with tariff", "wtariff", "terrif"],
#     }
#     keywords = hints.get(key, [])
#     lower_cols = [c.lower() for c in cols[1:]]
#     for kw in keywords:
#         for i, lc in enumerate(lower_cols):
#             if kw in lc:
#                 return i + 1
#     return 0


# def px(hex_color: str) -> PatternFill:
#     return PatternFill("solid", fgColor=hex_color)


# def thin_border() -> Border:
#     s = Side(style="thin", color="D0D0D0")
#     return Border(left=s, right=s, top=s, bottom=s)


# # ─────────────────────────────────────────────────────────────────────────────
# #  CORE EXCEL BUILDER
# #
# #  Layout (matches Image 2 / source file exactly):
# #   Row 1        : Headers
# #   Rows 2..DS+n-1 : Data  (n = number of stones)
# #   SUM_ROW      : TOTAL
# #   AVG_ROW      : AVG / CARAT
# #   (blank rows)
# #   FINAL_ROW    : FINAL label + value
# #   RATE_ROW     : RATE  label + value (yellow, red text) ← used by col-Y formula
# #   DIFF_ROW     : Difference
# #
# #  Column formulas (P-Y) replicate source exactly, adjusted for output row refs:
# #   P  =ROUND(((L-(M_AVG - O_AVG))/1.08), 2)      ← where M_AVG=M$AVG_ROW, O_AVG=O$AVG_ROW
# #   Q  =ROUND(P*D, 2)
# #   R  =ROUND(P*1.08, 2)
# #   S  =ROUND(L-R, 2)
# #   T  =ROUND((S*100)/P, 2)
# #   U  =ROUND((L*$I$SUM_ROW)/SUMPRODUCT($L$DS:$L$DE,$D$DS:$D$DE), 2)
# #   V  =ROUND(U*D, 2)
# #   W  =ROUND(L-U, 2)
# #   X  =ROUND((W*100)/U, 2)
# #   Y  =ROUND(U*$N$RATE_ROW, 2)   ← references RATE cell instead of hardcoded 95.95
# #
# #  SUM row  : D, I, K, M, O, Q, V
# #  AVG row  : I, K, M, O, Q, V  (each SUM/D_SUM)
# # ─────────────────────────────────────────────────────────────────────────────
# def build_output_excel(mapped_df: pd.DataFrame, rate_value: float = 95.95) -> bytes:
#     wb = openpyxl.Workbook()
#     ws = wb.active
#     ws.title = "Diamond Data"

#     n        = len(mapped_df)
#     HDR      = 1
#     DS       = 2            # data start row
#     DE       = DS + n - 1   # data end row
#     SUM_ROW  = DE + 1
#     AVG_ROW  = DE + 2
#     FINAL_ROW = DE + 5
#     RATE_ROW  = DE + 6
#     DIFF_ROW  = DE + 7

#     # ── Header row ────────────────────────────────────────────────────────
#     for c, hdr in enumerate(OUTPUT_HEADERS, 1):
#         cell = ws.cell(row=HDR, column=c, value=hdr)
#         cell.fill      = px(CLR_HEADER)
#         cell.font      = Font(name="Arial", bold=True, color="FFFFFF", size=10)
#         cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
#         cell.border    = thin_border()
#     ws.row_dimensions[HDR].height = 36

#     # ── Data rows ─────────────────────────────────────────────────────────
#     for i, (_, row) in enumerate(mapped_df.iterrows()):
#         r = DS + i

#         # Raw input columns (A-O in output = cols 1-15)
#         raw = {
#             1:  row.get("stock_no"),
#             2:  row.get("certi_no"),
#             3:  row.get("shape"),
#             4:  row.get("carat"),
#             5:  row.get("lab"),
#             6:  row.get("color"),
#             7:  row.get("clarity"),
#             8:  row.get("final"),       # H  Final (%)
#             10: row.get("vdb"),         # J  VDB (%)
#             12: row.get("usa"),         # L  USA (%)
#             14: row.get("with_tariff"), # N  With Tariff (%)
#         }

#         # Formula columns — EXACTLY matching source file formulas, re-rowed
#         # I9  = Final Amt   = D*H
#         # K11 = VDB Amt     = J*D
#         # M13 = USA Amt     = L*D
#         # O15 = Tariff Amt  = N*D   (source: N*1.08 then *D — actually source col14=H*1.08, col15=N*D)
#         #   wait — source col14 formula: =ROUND(H2*1.08,2)  so N=WithTariff col is H*1.08
#         #   and source col15: =ROUND(N2*D2,2)
#         #   In our output col14=WithTariff input, col15=TariffAmt
#         #   But source H=Final(input), N=H*1.08 (computed), O=N*D
#         #   Our mapping: col8=Final(input), col14=WithTariff(input from party)
#         #   To stay faithful: col9=D*H, col11=J*D, col13=L*D, col14=H*1.08, col15=N*D
#         #   BUT party file already has WithTariff as a value — so col14 = input from party
#         #   and col15 = N*D where N=col14
#         formulas = {
#             9:  f"=ROUND(D{r}*H{r},2)",           # I  Final Amt
#             11: f"=ROUND(J{r}*D{r},2)",            # K  VDB Amt
#             13: f"=ROUND(L{r}*D{r},2)",            # M  USA Amt
#             15: f"=ROUND(N{r}*D{r},2)",            # O  Tariff Amt

#             # ── YELLOW columns (P-Y) — exact source formulas ──
#             # P: =ROUND(((L-(M_AVG-O_AVG))/1.08),2)
#             16: f"=ROUND(((L{r}-(M${AVG_ROW}-O${AVG_ROW}))/1.08),2)",
#             # Q: =ROUND(P*D,2)
#             17: f"=ROUND(P{r}*D{r},2)",
#             # R: =ROUND(P*1.08,2)
#             18: f"=ROUND(P{r}*1.08,2)",
#             # S: =ROUND(L-R,2)
#             19: f"=ROUND(L{r}-R{r},2)",
#             # T: =ROUND((S*100)/P,2)
#             20: f"=ROUND((S{r}*100)/P{r},2)",
#             # U: =ROUND((L*$I$SUM)/SUMPRODUCT($L$DS:$L$DE,$D$DS:$D$DE),2)
#             21: f"=ROUND((L{r}*$I${SUM_ROW})/SUMPRODUCT($L${DS}:$L${DE},$D${DS}:$D${DE}),2)",
#             # V: =ROUND(U*D,2)
#             22: f"=ROUND(U{r}*D{r},2)",
#             # W: =ROUND(L-U,2)
#             23: f"=ROUND(L{r}-U{r},2)",
#             # X: =ROUND((W*100)/U,2)
#             24: f"=ROUND((W{r}*100)/U{r},2)",
#             # Y: =ROUND(U*RATE,2)  — references the RATE cell
#             25: f"=ROUND(U{r}*$N${RATE_ROW},2)",
#             # Z: =ROUND(Y*D,2)  — Total INR Amt
#             26: f"=ROUND(Y{r}*D{r},2)",
#         }

#         for c in range(1, 27):
#             val = raw.get(c, formulas.get(c))
#             cell = ws.cell(row=r, column=c, value=val)
#             cell.border    = thin_border()
#             cell.alignment = Alignment(horizontal="center", vertical="center")
#             cell.font      = Font(name="Arial", size=10)
#             if c in COL_COLORS:
#                 cell.fill = px(COL_COLORS[c])
#             if c in (4, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26):
#                 cell.number_format = "0.00"

#         ws.row_dimensions[r].height = 18

#     # ── TOTAL row ─────────────────────────────────────────────────────────
#     total_fill = px("F2F2F2")
#     lbl = ws.cell(row=SUM_ROW, column=1, value="TOTAL")
#     lbl.font = Font(name="Arial", bold=True, size=10)
#     lbl.fill = total_fill; lbl.border = thin_border()
#     lbl.alignment = Alignment(horizontal="center")

#     # SUM cols: D(4), I(9), K(11), M(13), O(15), Q(17), V(22), Z(26)
#     sum_cols = {4:"D", 9:"I", 11:"K", 13:"M", 15:"O", 17:"Q", 22:"V", 26:"Z"}
#     for col, letter in sum_cols.items():
#         cell = ws.cell(row=SUM_ROW, column=col,
#                        value=f"=SUM({letter}{DS}:{letter}{DE})")
#         cell.font = Font(name="Arial", bold=True, size=10)
#         cell.fill = total_fill; cell.border = thin_border()
#         cell.alignment = Alignment(horizontal="center")
#         cell.number_format = "0.00"

#     # ── AVG / CARAT row ───────────────────────────────────────────────────
#     avg_fill = px("E2EFDA")
#     lbl2 = ws.cell(row=AVG_ROW, column=1, value="AVG / CARAT")
#     lbl2.font = Font(name="Arial", bold=True, size=10)
#     lbl2.fill = avg_fill; lbl2.border = thin_border()
#     lbl2.alignment = Alignment(horizontal="center")

#     # AVG cols: I(9), K(11), M(13), O(15), Q(17), V(22)
#     avg_cols = {9:"I", 11:"K", 13:"M", 15:"O", 17:"Q", 22:"V"}
#     for col, letter in avg_cols.items():
#         cell = ws.cell(row=AVG_ROW, column=col,
#                        value=f"=ROUND({letter}{SUM_ROW}/D{SUM_ROW},2)")
#         cell.font = Font(name="Arial", bold=True, size=10)
#         cell.fill = avg_fill; cell.border = thin_border()
#         cell.alignment = Alignment(horizontal="center")
#         cell.number_format = "0.00"

#     # ── FINAL / RATE / Difference label block ────────────────────────────
#     # FINAL row
#     for c, val in ((13, "FINAL"), (14, f"=ROUND(I{AVG_ROW},2)")):
#         cell = ws.cell(row=FINAL_ROW, column=c, value=val)
#         cell.font = Font(name="Arial", bold=True, size=10)
#         cell.border = thin_border()
#         cell.alignment = Alignment(horizontal="center")
#         cell.number_format = "0.00"

#     # RATE row — label M, value N (yellow fill, red bold text for value)
#     lbl_rate = ws.cell(row=RATE_ROW, column=13, value="$ RATE")
#     lbl_rate.font = Font(name="Arial", bold=True, size=10)
#     lbl_rate.border = thin_border()
#     lbl_rate.alignment = Alignment(horizontal="center")

#     rate_cell = ws.cell(row=RATE_ROW, column=14, value=rate_value)
#     rate_cell.fill   = px("FFFF00")
#     rate_cell.font   = Font(name="Arial", bold=True, color="FF0000", size=10)
#     rate_cell.border = thin_border()
#     rate_cell.alignment = Alignment(horizontal="center")
#     rate_cell.number_format = "0.00"

#     # Difference row
#     for c, val in ((13, "Difference"), (14, f"=M{AVG_ROW}-O{AVG_ROW}")):
#         cell = ws.cell(row=DIFF_ROW, column=c, value=val)
#         cell.font = Font(name="Arial", bold=True, size=10)
#         cell.border = thin_border()
#         cell.alignment = Alignment(horizontal="center")
#         cell.number_format = "0.00"

#     # ── Column widths ──────────────────────────────────────────────────────
#     widths = [14, 14, 13, 7, 6, 7, 9,  8, 11,  8, 11,  8, 11, 12, 11,
#               11, 11, 11,  8,  9, 11, 11, 11,  9, 13, 15]
#     for i, w in enumerate(widths, 1):
#         ws.column_dimensions[get_column_letter(i)].width = w

#     ws.freeze_panes = "A2"

#     buf = io.BytesIO()
#     wb.save(buf)
#     buf.seek(0)
#     return buf.getvalue()


# # ─────────────────────────────────────────────────────────────────────────────
# #  HERO BANNER
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="hero">
#   <div class="hero-content">
#     <p class="hero-company">✦ BRIGHT TRADING INDIA LLP</p>
#     <p class="hero-title">
#       <span class="hero-gem-wrap">
#         <span class="gem">💎</span>
#         <span class="spark spark-1">✦</span>
#         <span class="spark spark-2">✦</span>
#       </span>
#       <span class="txt">Diamond Cost Optimizer</span>
#     </p>
#     <p class="hero-sub">Smart Cost Adjustment &amp; Equal Profit Calculator</p>
#   </div>
# </div>
# """, unsafe_allow_html=True)

# st.markdown("""
# <div class="legend-wrap">
#   <div class="pill"><div class="dot" style="background:#98F6EB"></div>Carat</div>
#   <div class="pill"><div class="dot" style="background:#FFCF37"></div>Final / Amt</div>
#   <div class="pill"><div class="dot" style="background:#5991D5"></div>VDB / Amt</div>
#   <div class="pill"><div class="dot" style="background:#F79B4F"></div>USA / Amt</div>
#   <div class="pill"><div class="dot" style="background:#9E89B9"></div>With Tariff / Amt</div>
#   <div class="pill"><div class="dot" style="background:#FFFF00"></div>Calculated Columns (P–Y)</div>
# </div>
# """, unsafe_allow_html=True)

# # ─────────────────────────────────────────────────────────────────────────────
# #  STEP 1 – UPLOAD
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="step-card">
#   <div class="step-header">
#     <span class="step-num">1</span>
#     <span class="step-title">Upload Party File</span>
#   </div>
#   <p style="color:#64748b;font-size:.87rem;margin:0 0 .9rem 42px;">
#     Accepts <b style="color:#a78bfa">CSV, XLSX, XLS, XLSM, TSV</b> — any column order, any layout.
#   </p>
# </div>
# """, unsafe_allow_html=True)

# party_file = st.file_uploader(
#     "Drop your file here or click Upload",
#     type=["csv", "xlsx", "xls", "xlsm", "tsv"],
#     key="party_file",
#     label_visibility="collapsed",
# )

# if party_file is None:
#     st.markdown('<div class="banner banner-info">👆 Upload the party file above to begin.</div>',
#                 unsafe_allow_html=True)
#     st.stop()

# raw_df = read_uploaded_file(party_file)
# if raw_df is None:
#     st.stop()

# h_row = detect_header_row(raw_df)
# df = raw_df.copy()
# df.columns = raw_df.iloc[h_row]
# df = df.iloc[h_row + 1:].reset_index(drop=True)
# df.columns = [str(c).strip() if pd.notna(c) else f"Col_{i}" for i, c in enumerate(df.columns)]
# df = df.dropna(how="all")

# # ── Auto-detect and strip SUM/footer rows at the bottom ──────────────────────
# # We do a quick keyword-based column sniff so trim works before Step-2 mapping
# def _sniff_col(df, keywords):
#     for col in df.columns:
#         if any(kw in str(col).lower() for kw in keywords):
#             return col
#     return None

# _stock_col = _sniff_col(df, ["stock", "packet", "pkt", "lot", "id"])
# _carat_col = _sniff_col(df, ["carat", "ct", "weight"])
# df = trim_data_rows(df, _stock_col, _carat_col)

# st.markdown(f"""
# <div class="stats-row">
#   <div class="stat-box"><div class="stat-val">{len(df)}</div><div class="stat-label">Rows Detected</div></div>
#   <div class="stat-box"><div class="stat-val">{len(df.columns)}</div><div class="stat-label">Columns Found</div></div>
#   <div class="stat-box"><div class="stat-val">{party_file.name.rsplit(".",1)[-1].upper()}</div><div class="stat-label">File Format</div></div>
# </div>
# """, unsafe_allow_html=True)

# with st.expander("👀 Preview raw data (first 5 rows)"):
#     st.dataframe(df.head(5), use_container_width=True)

# # ─────────────────────────────────────────────────────────────────────────────
# #  STEP 2 – COLUMN MAPPING
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="step-card">
#   <div class="step-header">
#     <span class="step-num">2</span>
#     <span class="step-title">Map Columns</span>
#   </div>
# </div>
# """, unsafe_allow_html=True)
# st.caption("Select which column in the party file corresponds to each required field.")

# available = ["— skip / not available —"] + list(df.columns)
# mapping   = {}

# cols_left, cols_right = st.columns(2)
# for idx, (key, label) in enumerate(REQUIRED_COLS):
#     default = smart_default(key, available)
#     container = cols_left if idx % 2 == 0 else cols_right
#     with container:
#         choice = st.selectbox(f"**{label}**", options=available,
#                               index=default, key=f"map_{key}")
#         mapping[key] = None if choice.startswith("—") else choice

# # ── $ RATE input ──────────────────────────────────────────────────────────────
# st.markdown("---")
# st.markdown("**$ Rate** — used in the final Y-column formula (`U × $ Rate`)")
# rate_input = st.number_input("$ Rate value", min_value=0.0, value=95.95, step=0.01, format="%.2f")

# mandatory = ["carat", "final", "vdb", "usa", "with_tariff"]
# missing   = [next(l for kk, l in REQUIRED_COLS if kk == k)
#              for k in mandatory if not mapping.get(k)]
# if missing:
#     st.markdown(f'<div class="banner banner-warn">⚠️ Please map required columns: <b>{", ".join(missing)}</b></div>',
#                 unsafe_allow_html=True)
#     st.stop()

# # ─────────────────────────────────────────────────────────────────────────────
# #  STEP 3 – PREVIEW & DOWNLOAD
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="step-card">
#   <div class="step-header">
#     <span class="step-num">3</span>
#     <span class="step-title">Preview &amp; Download</span>
#   </div>
# </div>
# """, unsafe_allow_html=True)

# # ── Re-trim using the user's actual mapped columns (catches any edge cases) ──
# df = trim_data_rows(df, mapping.get("stock_no"), mapping.get("carat"))

# records = []
# for _, row in df.iterrows():
#     rec = {key: (row[col] if col else None) for key, col in mapping.items()}
#     records.append(rec)

# mapped_df = pd.DataFrame(records)
# for col in ["carat", "final", "vdb", "usa", "with_tariff"]:
#     mapped_df[col] = pd.to_numeric(mapped_df[col], errors="coerce")

# mapped_df = mapped_df.dropna(subset=["carat", "final", "vdb", "usa", "with_tariff"], how="all")
# n_stones  = len(mapped_df)

# if n_stones == 0:
#     st.markdown('<div class="banner banner-warn">⚠️ No valid rows found. Check your column mappings.</div>',
#                 unsafe_allow_html=True)
#     st.stop()

# st.markdown(f'<div class="banner banner-success">✅ <b>{n_stones} stones</b> ready to export.</div>',
#             unsafe_allow_html=True)

# preview_rename = {
#     "stock_no":"Stock ID.", "certi_no":"Certi No.", "shape":"Shape",
#     "carat":"Carat", "lab":"Lab", "color":"Color", "clarity":"Clarity",
#     "final":"Final", "vdb":"VDB", "usa":"USA", "with_tariff":"With Tariff",
# }
# st.dataframe(mapped_df.rename(columns=preview_rename).head(10), use_container_width=True)

# with st.spinner("Building Excel with formulas and colours…"):
#     xlsx_bytes   = build_output_excel(mapped_df, rate_value=rate_input)
#     out_filename = safe_filename(party_file.name, n_stones)

# st.markdown('<div class="banner banner-success">💾 File ready — all formulas match the reference format exactly (cols P–Y yellow).</div>',
#             unsafe_allow_html=True)

# st.download_button(
#     label=f"⬇️  Download  {out_filename}",
#     data=xlsx_bytes,
#     file_name=out_filename,
#     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#     use_container_width=True,
# )

# with st.expander("📐 Formula & Colour Reference"):
#     st.markdown(f"""
# | Col | Header | Colour | Formula / Source |
# |---|---|---|---|
# | A | Stock ID | — | From party file |
# | B | Certi No | — | From party file |
# | C | Shape | — | From party file |
# | D | Carat | 🟦 Teal | From party file |
# | E | Lab | — | From party file |
# | F | Color | — | From party file |
# | G | Clarity | — | From party file |
# | H | Final | 🟨 Gold | From party file |
# | I | Amt | 🟨 Gold | `D × H` |
# | J | VDB | 🔵 Blue | From party file |
# | K | Amt | 🔵 Blue | `J × D` |
# | L | USA | 🟠 Orange | From party file |
# | M | Amt | 🟠 Orange | `L × D` |
# | N | With Tariff | 🟣 Lavender | From party file |
# | O | Amt | 🟣 Lavender | `N × D` |
# | P | — | 🟡 Yellow | `ROUND(((L-(M_avg−O_avg))/1.08),2)` |
# | Q | — | 🟡 Yellow | `ROUND(P×D, 2)` |
# | R | — | 🟡 Yellow | `ROUND(P×1.08, 2)` |
# | S | — | 🟡 Yellow | `ROUND(L−R, 2)` |
# | T | — | 🟡 Yellow | `ROUND((S×100)/P, 2)` |
# | U | — | 🟡 Yellow | `ROUND((L×ΣI) / SUMPRODUCT(ΣL,ΣD), 2)` |
# | V | — | 🟡 Yellow | `ROUND(U×D, 2)` |
# | W | — | 🟡 Yellow | `ROUND(L−U, 2)` |
# | X | — | 🟡 Yellow | `ROUND((W×100)/U, 2)` |
# | Y | — | 🟡 Yellow | `ROUND(U×$RATE, 2)` — RATE = **{rate_input}** |

# > 🟡 **$ RATE** cell is highlighted yellow/red in the summary block — editable directly in Excel.
# """)

# Working A TO Z 


# import streamlit as st
# import pandas as pd
# import openpyxl
# from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
# from openpyxl.utils import get_column_letter
# import io
# import re
# import os
# from datetime import datetime

# # ─────────────────────────────────────────────────────────────────────────────
# #  PAGE CONFIG
# # ─────────────────────────────────────────────────────────────────────────────
# st.set_page_config(
#     page_title="💎 Diamond Formatter",
#     page_icon="💎",
#     layout="wide",
#     initial_sidebar_state="collapsed",
# )

# # ─────────────────────────────────────────────────────────────────────────────
# #  DARK THEME CSS
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <style>
# html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"] {
#     background-color: #0d0f14 !important; color: #e2e8f0 !important;
# }
# [data-testid="stHeader"] { background: #0d0f14 !important; }
# [data-testid="stSidebar"] { background: #111318 !important; }
# section[data-testid="stMain"] > div { background: #0d0f14 !important; }
# .block-container { padding-top: 1.5rem !important; }
# .hero {
#     position: relative;
#     background: linear-gradient(135deg, #0f1629 0%, #1a0a2e 50%, #0a1628 100%);
#     background-size: 200% 200%;
#     animation: heroGradientShift 12s ease infinite;
#     border: 1px solid #2a2d3e; border-radius: 20px;
#     padding: 2.4rem 2rem; margin-bottom: 2rem; overflow: hidden;
#     text-align: center; box-shadow: 0 10px 40px rgba(0,0,0,.35);
# }
# @keyframes heroGradientShift {
#     0%   { background-position: 0% 50%; }
#     50%  { background-position: 100% 50%; }
#     100% { background-position: 0% 50%; }
# }
# .hero::before, .hero::after {
#     content: ''; position: absolute; border-radius: 50%;
#     filter: blur(60px); opacity: .55;
#     animation: heroFloat 8s ease-in-out infinite; pointer-events: none;
# }
# .hero::before { width:240px;height:240px;background:rgba(99,102,241,.35);top:-70px;left:8%; }
# .hero::after  { width:280px;height:280px;background:rgba(168,85,247,.28);bottom:-90px;right:8%;animation-delay:-4s; }
# @keyframes heroFloat {
#     0%,100% { transform:translateY(0) translateX(0); }
#     50%     { transform:translateY(-22px) translateX(16px); }
# }
# .hero-content { position:relative;z-index:1;text-align:left; }
# .hero-company {
#     display:inline-block;font-size:.92rem;font-weight:700;letter-spacing:2.6px;
#     color:#c4b5fd;text-transform:uppercase;padding:7px 20px;
#     border:1px solid rgba(167,139,250,.45);border-radius:30px;
#     background:rgba(99,102,241,.12);margin:0 0 1.1rem;
#     animation:heroFadeInDown .8s ease both,heroPulseBorder 3s ease-in-out infinite 1s;
# }
# @keyframes heroFadeInDown { from{opacity:0;transform:translateY(-14px)} to{opacity:1;transform:translateY(0)} }
# @keyframes heroPulseBorder {
#     0%,100% { box-shadow:0 0 0 rgba(167,139,250,0); }
#     50%     { box-shadow:0 0 20px rgba(167,139,250,.4); }
# }
# .hero-title { display:flex;align-items:center;justify-content:flex-start;gap:12px;margin:0 0 .6rem;animation:heroFadeInUp .9s ease .2s both; }
# .hero-gem-wrap { position:relative;display:inline-flex;align-items:center;justify-content:center; }
# .hero-title .gem { display:inline-block;font-size:2rem;line-height:1;transform-style:preserve-3d;animation:heroGemSpin 4.5s ease-in-out infinite; }
# .hero-gem-wrap .spark { position:absolute;color:#7dd3fc;font-size:.6rem;opacity:0; }
# .hero-gem-wrap .spark-1 { top:-6px;right:-8px;animation:heroTwinkle 2.6s ease-in-out infinite; }
# .hero-gem-wrap .spark-2 { bottom:-4px;left:-10px;animation:heroTwinkle 2.6s ease-in-out infinite 1.1s; }
# @keyframes heroTwinkle { 0%,100%{opacity:0;transform:scale(.4) rotate(0deg)} 50%{opacity:1;transform:scale(1.1) rotate(25deg)} }
# @keyframes heroGemSpin {
#     0%   { transform:perspective(320px) rotateY(0deg) scale(1);filter:drop-shadow(0 0 6px rgba(167,139,250,.45)); }
#     45%  { transform:perspective(320px) rotateY(180deg) scale(1.1);filter:drop-shadow(0 0 16px rgba(56,189,248,.8)); }
#     55%  { transform:perspective(320px) rotateY(180deg) scale(1.1);filter:drop-shadow(0 0 16px rgba(56,189,248,.8)); }
#     100% { transform:perspective(320px) rotateY(360deg) scale(1);filter:drop-shadow(0 0 6px rgba(167,139,250,.45)); }
# }
# .hero-title .txt {
#     font-size:2.3rem;font-weight:800;letter-spacing:-.5px;line-height:1.15;
#     background:linear-gradient(90deg,#a78bfa,#6366f1,#38bdf8,#a78bfa);
#     background-size:300% auto;
#     -webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;
#     animation:heroShimmer 6s linear infinite;
# }
# @keyframes heroShimmer { to{background-position:300% center} }
# @keyframes heroFadeInUp { from{opacity:0;transform:translateY(18px)} to{opacity:1;transform:translateY(0)} }
# .hero-sub { color:#94a3b8;font-size:1rem;margin:0;animation:heroFadeInUp .9s ease .45s both; }
# .legend-wrap { display:flex;flex-wrap:wrap;gap:8px;margin:1rem 0 1.5rem; }
# .pill {
#     display:inline-flex;align-items:center;gap:7px;
#     background:#1a1d27;border:1px solid #2a2d3e;
#     border-radius:20px;padding:5px 12px;font-size:.78rem;color:#cbd5e1;
#     transition:transform .18s ease,border-color .18s ease;
# }
# .pill:hover { transform:translateY(-2px);border-color:#6366f1; }
# .dot { width:12px;height:12px;border-radius:3px;flex-shrink:0; }
# .step-card { background:#13161f;border:1px solid #1e2130;border-radius:12px;padding:1.5rem;margin-bottom:1rem; }
# .step-num {
#     display:inline-flex;align-items:center;justify-content:center;
#     width:32px;height:32px;border-radius:50%;
#     background:linear-gradient(135deg,#6366f1,#8b5cf6);
#     color:white;font-weight:700;font-size:.9rem;margin-right:10px;flex-shrink:0;
# }
# .step-title { font-size:1.1rem;font-weight:600;color:#f1f5f9; }
# .step-header { display:flex;align-items:center;margin-bottom:1rem; }
# [data-testid="stFileUploader"] { background:#13161f !important;border:2px dashed #2a2d3e !important;border-radius:10px !important; }
# [data-testid="stFileUploader"]:hover { border-color:#6366f1 !important; }
# .stSelectbox > div > div { background:#1a1d27 !important;border:1px solid #2a2d3e !important;color:#e2e8f0 !important;border-radius:8px !important; }
# .stSelectbox label { color:#cbd5e1 !important;font-size:.85rem !important; }
# .stDownloadButton button {
#     background:linear-gradient(135deg,#6366f1,#8b5cf6) !important;
#     color:white !important;border:none !important;border-radius:10px !important;
#     font-weight:600 !important;padding:.75rem 1.5rem !important;
#     font-size:1rem !important;transition:all .2s !important;
#     box-shadow:0 4px 20px rgba(99,102,241,.4) !important;
# }
# .stDownloadButton button:hover { transform:translateY(-2px) !important;box-shadow:0 8px 30px rgba(99,102,241,.6) !important; }
# .banner { border-radius:10px;padding:.85rem 1.1rem;font-size:.9rem;margin:.5rem 0;display:flex;align-items:center;gap:10px; }
# .banner-success { background:rgba(34,197,94,.1);border:1px solid rgba(34,197,94,.3);color:#86efac; }
# .banner-info    { background:rgba(99,102,241,.1);border:1px solid rgba(99,102,241,.3);color:#a5b4fc; }
# .banner-warn    { background:rgba(234,179,8,.1); border:1px solid rgba(234,179,8,.3); color:#fde68a; }
# .stats-row { display:flex;gap:12px;flex-wrap:wrap;margin:1rem 0; }
# .stat-box { background:#13161f;border:1px solid #1e2130;border-radius:10px;padding:.75rem 1.25rem;flex:1;min-width:130px; }
# .stat-val { font-size:1.6rem;font-weight:700;color:#a78bfa; }
# .stat-label { font-size:.75rem;color:#64748b;text-transform:uppercase;letter-spacing:.5px; }
# hr { border-color:#1e2130 !important;margin:1.5rem 0 !important; }
# [data-testid="stExpander"] { background:#13161f !important;border:1px solid #1e2130 !important;border-radius:10px !important; }
# summary { color:#94a3b8 !important; }
# </style>
# """, unsafe_allow_html=True)

# # ─────────────────────────────────────────────────────────────────────────────
# #  COLOR CONSTANTS
# #  Matches Image 2 exactly:
# #  - Cols A-O : existing group colors (navy header, teal, gold, blue, orange, lavender)
# #  - Cols P-Y : YELLOW (FFFF00) — boss requirement
# # ─────────────────────────────────────────────────────────────────────────────
# CLR_HEADER  = "1F4E79"   # deep navy header
# CLR_CARAT   = "98F6EB"   # teal   → col D (4)
# CLR_FINAL   = "FFCF37"   # gold   → H,I (8,9)
# CLR_VDB     = "5991D5"   # blue   → J,K (10,11)
# CLR_USA     = "F79B4F"   # orange → L,M (12,13)
# CLR_TARIFF  = "9E89B9"   # lavender → N,O (14,15)
# CLR_YELLOW  = "FFFF00"   # yellow → P-Y (16-25)

# # col index (1-based) → fill color for DATA cells
# COL_COLORS = {
#     4:  CLR_CARAT,
#     8:  CLR_FINAL,   9:  CLR_FINAL,
#     10: CLR_VDB,     11: CLR_VDB,
#     12: CLR_USA,     13: CLR_USA,
#     14: CLR_TARIFF,  15: CLR_TARIFF,
# }
# # cols 16-26 all yellow
# for _c in range(16, 27):
#     COL_COLORS[_c] = CLR_YELLOW

# # ─────────────────────────────────────────────────────────────────────────────
# #  OUTPUT HEADERS  (25 columns, matching Image 2)
# # ─────────────────────────────────────────────────────────────────────────────
# OUTPUT_HEADERS = [
#     "Stock ID.",        # A  1
#     "Certi No.",        # B  2
#     "Shape",            # C  3
#     "Carat",            # D  4
#     "Lab",              # E  5
#     "Color",            # F  6
#     "Clarity",          # G  7
#     "Final",            # H  8
#     "Amt",              # I  9
#     "VDB",              # J  10
#     "Amt",              # K  11
#     "USA",              # L  12
#     "Amt",              # M  13
#     "With Tariff",      # N  14
#     "Amt",              # O  15
#     "",                 # P  16  (yellow — unlabeled in source)
#     "",                 # Q  17
#     "",                 # R  18
#     "",                 # S  19
#     "",                 # T  20
#     "",                 # U  21
#     "",                 # V  22
#     "",                 # W  23
#     "",                 # X  24
#     "",                 # Y  25
#     "Total INR Amt",   # Z  26
# ]

# # NOTE: "with_tariff" is no longer taken from the party file — it is now
# # CALCULATED as Final × Tariff-Multiplier (user enters the multiplier, e.g.
# # 1.08 or 1.06), so it has been removed from the columns the user maps.
# REQUIRED_COLS = [
#     ("stock_no",    "Stock ID."),
#     ("certi_no",    "Certi No."),
#     ("shape",       "Shape"),
#     ("carat",       "Carat"),
#     ("lab",         "Lab"),
#     ("color",       "Color"),
#     ("clarity",     "Clarity"),
#     ("final",       "Final"),
#     ("vdb",         "VDB"),
#     ("usa",         "USA"),
# ]

# # ─────────────────────────────────────────────────────────────────────────────
# #  HELPERS
# # ─────────────────────────────────────────────────────────────────────────────
# def read_uploaded_file(f) -> pd.DataFrame | None:
#     name = f.name.lower()
#     try:
#         if name.endswith(".csv"):
#             return pd.read_csv(f)
#         elif name.endswith((".xlsx", ".xls", ".xlsm")):
#             return pd.read_excel(f, header=None)
#         elif name.endswith(".tsv"):
#             return pd.read_csv(f, sep="\t")
#         else:
#             st.error("Unsupported format. Please upload CSV, XLSX, XLS, XLSM, or TSV.")
#             return None
#     except Exception as e:
#         st.error(f"Could not read file: {e}")
#         return None


# def detect_header_row(df: pd.DataFrame) -> int:
#     for i, row in df.iterrows():
#         vals = [str(v).strip() for v in row if pd.notna(v) and str(v).strip()]
#         if len(vals) >= 4 and not all(
#             re.sub(r'[.\-]', '', v).isdigit() for v in vals
#         ):
#             return i
#     return 0


# def trim_data_rows(df: pd.DataFrame, stock_col: str | None, carat_col: str | None) -> pd.DataFrame:
#     """
#     Drop any trailing non-data rows (SUM totals, labels, blanks) that appear
#     after the real diamond records.

#     A row is treated as a SUM / footer row — and all rows from it onward are
#     dropped — when ANY of these conditions holds:

#     1. Carat column is mapped AND the carat cell has a numeric value BUT the
#        stock-ID cell is empty or NaN.  This is the classic pattern:
#            Stock  | Carat
#            (blank)| 28.41   ← SUM row
#     2. Carat column is mapped AND the carat value is ≥ 3× the median carat of
#        the rows seen so far (catches a huge SUM value even if stock is filled).
#     3. Stock-ID column is mapped AND the stock-ID cell is a purely numeric string
#        or a known summary keyword (TOTAL, AVG, SUM, RATE, FINAL, etc.).
#     """
#     if df.empty:
#         return df

#     summary_keywords = {"total", "avg", "average", "sum", "rate", "final",
#                         "difference", "subtotal", "done", "fianl"}

#     # Compute median carat from first 5 valid rows for outlier detection
#     median_carat = None
#     if carat_col and carat_col in df.columns:
#         sample_carats = pd.to_numeric(df[carat_col].head(20), errors="coerce").dropna()
#         if not sample_carats.empty:
#             median_carat = sample_carats.median()

#     cut_at = len(df)  # default: keep everything

#     for i, (_, row) in enumerate(df.iterrows()):
#         # ── condition 1: empty stock + non-empty carat ──
#         if stock_col and stock_col in df.columns and carat_col and carat_col in df.columns:
#             stock_val = row.get(stock_col)
#             carat_val = row.get(carat_col)
#             stock_empty = pd.isna(stock_val) or str(stock_val).strip() == ""
#             carat_numeric = pd.notna(carat_val) and str(carat_val).strip() != ""
#             if stock_empty and carat_numeric:
#                 cut_at = i
#                 break

#         # ── condition 2: carat value is a huge outlier (SUM) ──
#         if carat_col and carat_col in df.columns and median_carat:
#             carat_val = pd.to_numeric(row.get(carat_col), errors="coerce")
#             if pd.notna(carat_val) and carat_val >= median_carat * 3:
#                 cut_at = i
#                 break

#         # ── condition 3: stock ID is a summary keyword ──
#         if stock_col and stock_col in df.columns:
#             stock_val = str(row.get(stock_col, "")).strip().lower()
#             if stock_val in summary_keywords:
#                 cut_at = i
#                 break

#     return df.iloc[:cut_at].reset_index(drop=True)


# def safe_filename(raw_name: str, n_stones: int) -> str:
#     """Party_Name_No_of_Stones_Current_Date.xlsx"""
#     stem = os.path.splitext(raw_name)[0]
#     stem = re.sub(r'[^\w\s\-]', '', stem).strip()
#     stem = re.sub(r'\s+', '_', stem)
#     date_str = datetime.now().strftime("%d-%m-%Y")
#     return f"{stem}_{n_stones}_Stones_{date_str}.xlsx"


# def smart_default(key: str, cols: list) -> int:
#     hints = {
#         "stock_no":    ["stock", "pexkt", "packet", "pkt", "id", "lot"],
#         "certi_no":    ["certi", "cert", "gia", "igi", "report"],
#         "shape":       ["shape"],
#         "carat":       ["carat", "ct", "weight"],
#         "lab":         ["lab", "igi", "gia"],
#         "color":       ["color", "colour"],
#         "clarity":     ["clarity", "clar"],
#         "final":       ["final", "rap%", "disc%", "back%"],
#         "vdb":         ["vdb"],
#         "usa":         ["usa"],
#     }
#     keywords = hints.get(key, [])
#     lower_cols = [c.lower() for c in cols[1:]]
#     for kw in keywords:
#         for i, lc in enumerate(lower_cols):
#             if kw in lc:
#                 return i + 1
#     return 0


# def px(hex_color: str) -> PatternFill:
#     return PatternFill("solid", fgColor=hex_color)


# def thin_border() -> Border:
#     s = Side(style="thin", color="D0D0D0")
#     return Border(left=s, right=s, top=s, bottom=s)


# # ─────────────────────────────────────────────────────────────────────────────
# #  CORE EXCEL BUILDER
# #
# #  Layout (matches Image 2 / source file exactly):
# #   Row 1        : Headers
# #   Rows 2..DS+n-1 : Data  (n = number of stones)
# #   SUM_ROW      : TOTAL
# #   AVG_ROW      : AVG / CARAT
# #   (blank rows)
# #   FINAL_ROW    : FINAL label + value
# #   TARIFF_ROW   : TARIFF x  label + value (yellow, red text) ← used by col-N formula
# #   RATE_ROW     : $ RATE label + value (yellow, red text) ← used by col-Y formula
# #   DIFF_ROW     : Difference
# #
# #  Column formulas (H onward) replicate source exactly, adjusted for output row refs:
# #   N  =ROUND(H*$N$TARIFF_ROW, 2)                 ← With Tariff = Final × Tariff multiplier
# #   O  =ROUND(N*D, 2)
# #   P  =ROUND(((L-(M_AVG - O_AVG))/1.08), 2)      ← where M_AVG=M$AVG_ROW, O_AVG=O$AVG_ROW
# #   Q  =ROUND(P*D, 2)
# #   R  =ROUND(P*1.08, 2)
# #   S  =ROUND(L-R, 2)
# #   T  =ROUND((S*100)/P, 2)
# #   U  =ROUND((L*$I$SUM_ROW)/SUMPRODUCT($L$DS:$L$DE,$D$DS:$D$DE), 2)
# #   V  =ROUND(U*D, 2)
# #   W  =ROUND(L-U, 2)
# #   X  =ROUND((W*100)/U, 2)
# #   Y  =ROUND(U*$N$RATE_ROW, 2)   ← references RATE cell instead of hardcoded 95.95
# #
# #  SUM row  : D, I, K, M, O, Q, V
# #  AVG row  : I, K, M, O, Q, V  (each SUM/D_SUM)
# # ─────────────────────────────────────────────────────────────────────────────
# def build_output_excel(mapped_df: pd.DataFrame, rate_value: float = 95.95,
#                         tariff_value: float = 1.08) -> bytes:
#     wb = openpyxl.Workbook()
#     ws = wb.active
#     ws.title = "Diamond Data"

#     n        = len(mapped_df)
#     HDR      = 1
#     DS       = 2            # data start row
#     DE       = DS + n - 1   # data end row
#     SUM_ROW    = DE + 1
#     AVG_ROW    = DE + 2
#     FINAL_ROW  = DE + 5
#     TARIFF_ROW = DE + 6
#     RATE_ROW   = DE + 7
#     DIFF_ROW   = DE + 8

#     # ── Header row ────────────────────────────────────────────────────────
#     for c, hdr in enumerate(OUTPUT_HEADERS, 1):
#         cell = ws.cell(row=HDR, column=c, value=hdr)
#         cell.fill      = px(CLR_HEADER)
#         cell.font      = Font(name="Arial", bold=True, color="FFFFFF", size=10)
#         cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
#         cell.border    = thin_border()
#     ws.row_dimensions[HDR].height = 36

#     # ── Data rows ─────────────────────────────────────────────────────────
#     for i, (_, row) in enumerate(mapped_df.iterrows()):
#         r = DS + i

#         # Raw input columns (A-M in output = cols 1-13, from party file)
#         raw = {
#             1:  row.get("stock_no"),
#             2:  row.get("certi_no"),
#             3:  row.get("shape"),
#             4:  row.get("carat"),
#             5:  row.get("lab"),
#             6:  row.get("color"),
#             7:  row.get("clarity"),
#             8:  row.get("final"),       # H  Final (%)
#             10: row.get("vdb"),         # J  VDB (%)
#             12: row.get("usa"),         # L  USA (%)
#         }

#         # Formula columns — col N (14) is now CALCULATED as Final × Tariff
#         # multiplier instead of being read from the party file.
#         formulas = {
#             9:  f"=ROUND(D{r}*H{r},2)",           # I  Final Amt
#             11: f"=ROUND(J{r}*D{r},2)",            # K  VDB Amt
#             13: f"=ROUND(L{r}*D{r},2)",            # M  USA Amt

#             # N  With Tariff = Final × Tariff multiplier (user-entered, cell ref)
#             14: f"=ROUND(H{r}*$N${TARIFF_ROW},2)",
#             # O  Tariff Amt = N*D
#             15: f"=ROUND(N{r}*D{r},2)",

#             # ── YELLOW columns (P-Y) — exact source formulas ──
#             # P: =ROUND(((L-(M_AVG-O_AVG))/1.08),2)
#             16: f"=ROUND(((L{r}-(M${AVG_ROW}-O${AVG_ROW}))/1.08),2)",
#             # Q: =ROUND(P*D,2)
#             17: f"=ROUND(P{r}*D{r},2)",
#             # R: =ROUND(P*1.08,2)
#             18: f"=ROUND(P{r}*1.08,2)",
#             # S: =ROUND(L-R,2)
#             19: f"=ROUND(L{r}-R{r},2)",
#             # T: =ROUND((S*100)/P,2)
#             20: f"=ROUND((S{r}*100)/P{r},2)",
#             # U: =ROUND((L*$I$SUM)/SUMPRODUCT($L$DS:$L$DE,$D$DS:$D$DE),2)
#             21: f"=ROUND((L{r}*$I${SUM_ROW})/SUMPRODUCT($L${DS}:$L${DE},$D${DS}:$D${DE}),2)",
#             # V: =ROUND(U*D,2)
#             22: f"=ROUND(U{r}*D{r},2)",
#             # W: =ROUND(L-U,2)
#             23: f"=ROUND(L{r}-U{r},2)",
#             # X: =ROUND((W*100)/U,2)
#             24: f"=ROUND((W{r}*100)/U{r},2)",
#             # Y: =ROUND(U*RATE,2)  — references the RATE cell
#             25: f"=ROUND(U{r}*$N${RATE_ROW},2)",
#             # Z: =ROUND(Y*D,2)  — Total INR Amt
#             26: f"=ROUND(Y{r}*D{r},2)",
#         }

#         for c in range(1, 27):
#             val = raw.get(c, formulas.get(c))
#             cell = ws.cell(row=r, column=c, value=val)
#             cell.border    = thin_border()
#             cell.alignment = Alignment(horizontal="center", vertical="center")
#             cell.font      = Font(name="Arial", size=10)
#             if c in COL_COLORS:
#                 cell.fill = px(COL_COLORS[c])
#             if c in (4, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26):
#                 cell.number_format = "0.00"

#         ws.row_dimensions[r].height = 18

#     # ── TOTAL row ─────────────────────────────────────────────────────────
#     total_fill = px("F2F2F2")
#     lbl = ws.cell(row=SUM_ROW, column=1, value="TOTAL")
#     lbl.font = Font(name="Arial", bold=True, size=10)
#     lbl.fill = total_fill; lbl.border = thin_border()
#     lbl.alignment = Alignment(horizontal="center")

#     # SUM cols: D(4), I(9), K(11), M(13), O(15), Q(17), V(22), Z(26)
#     sum_cols = {4:"D", 9:"I", 11:"K", 13:"M", 15:"O", 17:"Q", 22:"V", 26:"Z"}
#     for col, letter in sum_cols.items():
#         cell = ws.cell(row=SUM_ROW, column=col,
#                        value=f"=SUM({letter}{DS}:{letter}{DE})")
#         cell.font = Font(name="Arial", bold=True, size=10)
#         cell.fill = total_fill; cell.border = thin_border()
#         cell.alignment = Alignment(horizontal="center")
#         cell.number_format = "0.00"

#     # ── AVG / CARAT row ───────────────────────────────────────────────────
#     avg_fill = px("E2EFDA")
#     lbl2 = ws.cell(row=AVG_ROW, column=1, value="AVG / CARAT")
#     lbl2.font = Font(name="Arial", bold=True, size=10)
#     lbl2.fill = avg_fill; lbl2.border = thin_border()
#     lbl2.alignment = Alignment(horizontal="center")

#     # AVG cols: I(9), K(11), M(13), O(15), Q(17), V(22)
#     avg_cols = {9:"I", 11:"K", 13:"M", 15:"O", 17:"Q", 22:"V"}
#     for col, letter in avg_cols.items():
#         cell = ws.cell(row=AVG_ROW, column=col,
#                        value=f"=ROUND({letter}{SUM_ROW}/D{SUM_ROW},2)")
#         cell.font = Font(name="Arial", bold=True, size=10)
#         cell.fill = avg_fill; cell.border = thin_border()
#         cell.alignment = Alignment(horizontal="center")
#         cell.number_format = "0.00"

#     # ── FINAL / TARIFF / RATE / Difference label block ───────────────────
#     # FINAL row
#     for c, val in ((13, "FINAL"), (14, f"=ROUND(I{AVG_ROW},2)")):
#         cell = ws.cell(row=FINAL_ROW, column=c, value=val)
#         cell.font = Font(name="Arial", bold=True, size=10)
#         cell.border = thin_border()
#         cell.alignment = Alignment(horizontal="center")
#         cell.number_format = "0.00"

#     # TARIFF row — label M, value N (yellow fill, red bold text for value)
#     # This is the multiplier used to CALCULATE "With Tariff" (col N) as
#     # Final × Tariff, e.g. 1.08 or 1.06 — editable directly in Excel.
#     lbl_tariff = ws.cell(row=TARIFF_ROW, column=13, value="TARIFF x")
#     lbl_tariff.font = Font(name="Arial", bold=True, size=10)
#     lbl_tariff.border = thin_border()
#     lbl_tariff.alignment = Alignment(horizontal="center")

#     tariff_cell = ws.cell(row=TARIFF_ROW, column=14, value=tariff_value)
#     tariff_cell.fill   = px("FFFF00")
#     tariff_cell.font   = Font(name="Arial", bold=True, color="FF0000", size=10)
#     tariff_cell.border = thin_border()
#     tariff_cell.alignment = Alignment(horizontal="center")
#     tariff_cell.number_format = "0.00"

#     # RATE row — label M, value N (yellow fill, red bold text for value)
#     lbl_rate = ws.cell(row=RATE_ROW, column=13, value="$ RATE")
#     lbl_rate.font = Font(name="Arial", bold=True, size=10)
#     lbl_rate.border = thin_border()
#     lbl_rate.alignment = Alignment(horizontal="center")

#     rate_cell = ws.cell(row=RATE_ROW, column=14, value=rate_value)
#     rate_cell.fill   = px("FFFF00")
#     rate_cell.font   = Font(name="Arial", bold=True, color="FF0000", size=10)
#     rate_cell.border = thin_border()
#     rate_cell.alignment = Alignment(horizontal="center")
#     rate_cell.number_format = "0.00"

#     # Difference row
#     for c, val in ((13, "Difference"), (14, f"=M{AVG_ROW}-O{AVG_ROW}")):
#         cell = ws.cell(row=DIFF_ROW, column=c, value=val)
#         cell.font = Font(name="Arial", bold=True, size=10)
#         cell.border = thin_border()
#         cell.alignment = Alignment(horizontal="center")
#         cell.number_format = "0.00"

#     # ── Column widths ──────────────────────────────────────────────────────
#     widths = [14, 14, 13, 7, 6, 7, 9,  8, 11,  8, 11,  8, 11, 12, 11,
#               11, 11, 11,  8,  9, 11, 11, 11,  9, 13, 15]
#     for i, w in enumerate(widths, 1):
#         ws.column_dimensions[get_column_letter(i)].width = w

#     ws.freeze_panes = "A2"

#     buf = io.BytesIO()
#     wb.save(buf)
#     buf.seek(0)
#     return buf.getvalue()


# # ─────────────────────────────────────────────────────────────────────────────
# #  HERO BANNER
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="hero">
#   <div class="hero-content">
#     <p class="hero-company">✦ BRIGHT TRADING INDIA LLP</p>
#     <p class="hero-title">
#       <span class="hero-gem-wrap">
#         <span class="gem">💎</span>
#         <span class="spark spark-1">✦</span>
#         <span class="spark spark-2">✦</span>
#       </span>
#       <span class="txt">Diamond Cost Optimizer</span>
#     </p>
#     <p class="hero-sub">Smart Cost Adjustment &amp; Equal Profit Calculator</p>
#   </div>
# </div>
# """, unsafe_allow_html=True)

# st.markdown("""
# <div class="legend-wrap">
#   <div class="pill"><div class="dot" style="background:#98F6EB"></div>Carat</div>
#   <div class="pill"><div class="dot" style="background:#FFCF37"></div>Final / Amt</div>
#   <div class="pill"><div class="dot" style="background:#5991D5"></div>VDB / Amt</div>
#   <div class="pill"><div class="dot" style="background:#F79B4F"></div>USA / Amt</div>
#   <div class="pill"><div class="dot" style="background:#9E89B9"></div>With Tariff / Amt</div>
#   <div class="pill"><div class="dot" style="background:#FFFF00"></div>Calculated Columns (P–Y)</div>
# </div>
# """, unsafe_allow_html=True)

# # ─────────────────────────────────────────────────────────────────────────────
# #  STEP 1 – UPLOAD
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="step-card">
#   <div class="step-header">
#     <span class="step-num">1</span>
#     <span class="step-title">Upload Party File</span>
#   </div>
#   <p style="color:#64748b;font-size:.87rem;margin:0 0 .9rem 42px;">
#     Accepts <b style="color:#a78bfa">CSV, XLSX, XLS, XLSM, TSV</b> — any column order, any layout.
#   </p>
# </div>
# """, unsafe_allow_html=True)

# party_file = st.file_uploader(
#     "Drop your file here or click Upload",
#     type=["csv", "xlsx", "xls", "xlsm", "tsv"],
#     key="party_file",
#     label_visibility="collapsed",
# )

# if party_file is None:
#     st.markdown('<div class="banner banner-info">👆 Upload the party file above to begin.</div>',
#                 unsafe_allow_html=True)
#     st.stop()

# raw_df = read_uploaded_file(party_file)
# if raw_df is None:
#     st.stop()

# h_row = detect_header_row(raw_df)
# df = raw_df.copy()
# df.columns = raw_df.iloc[h_row]
# df = df.iloc[h_row + 1:].reset_index(drop=True)
# df.columns = [str(c).strip() if pd.notna(c) else f"Col_{i}" for i, c in enumerate(df.columns)]
# df = df.dropna(how="all")

# # ── Auto-detect and strip SUM/footer rows at the bottom ──────────────────────
# # We do a quick keyword-based column sniff so trim works before Step-2 mapping
# def _sniff_col(df, keywords):
#     for col in df.columns:
#         if any(kw in str(col).lower() for kw in keywords):
#             return col
#     return None

# _stock_col = _sniff_col(df, ["stock", "packet", "pkt", "lot", "id"])
# _carat_col = _sniff_col(df, ["carat", "ct", "weight"])
# df = trim_data_rows(df, _stock_col, _carat_col)

# st.markdown(f"""
# <div class="stats-row">
#   <div class="stat-box"><div class="stat-val">{len(df)}</div><div class="stat-label">Rows Detected</div></div>
#   <div class="stat-box"><div class="stat-val">{len(df.columns)}</div><div class="stat-label">Columns Found</div></div>
#   <div class="stat-box"><div class="stat-val">{party_file.name.rsplit(".",1)[-1].upper()}</div><div class="stat-label">File Format</div></div>
# </div>
# """, unsafe_allow_html=True)

# with st.expander("👀 Preview raw data (first 5 rows)"):
#     st.dataframe(df.head(5), use_container_width=True)

# # ─────────────────────────────────────────────────────────────────────────────
# #  STEP 2 – COLUMN MAPPING
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="step-card">
#   <div class="step-header">
#     <span class="step-num">2</span>
#     <span class="step-title">Map Columns</span>
#   </div>
# </div>
# """, unsafe_allow_html=True)
# st.caption("Select which column in the party file corresponds to each required field.")

# available = ["— skip / not available —"] + list(df.columns)
# mapping   = {}

# cols_left, cols_right = st.columns(2)
# for idx, (key, label) in enumerate(REQUIRED_COLS):
#     default = smart_default(key, available)
#     container = cols_left if idx % 2 == 0 else cols_right
#     with container:
#         choice = st.selectbox(f"**{label}**", options=available,
#                               index=default, key=f"map_{key}")
#         mapping[key] = None if choice.startswith("—") else choice

# # ── Tariff multiplier + $ RATE inputs ─────────────────────────────────────────
# st.markdown("---")
# st.markdown("**Tariff Multiplier** — used to calculate the `With Tariff` column (`Final × Multiplier`), e.g. 1.08 or 1.06")
# tariff_input = st.number_input("Tariff multiplier", min_value=1.0, value=1.08, step=0.01, format="%.2f")

# st.markdown("**$ Rate** — used in the final Y-column formula (`U × $ Rate`)")
# rate_input = st.number_input("$ Rate value", min_value=0.0, value=95.95, step=0.01, format="%.2f")

# mandatory = ["carat", "final", "vdb", "usa"]
# missing   = [next(l for kk, l in REQUIRED_COLS if kk == k)
#              for k in mandatory if not mapping.get(k)]
# if missing:
#     st.markdown(f'<div class="banner banner-warn">⚠️ Please map required columns: <b>{", ".join(missing)}</b></div>',
#                 unsafe_allow_html=True)
#     st.stop()

# # ─────────────────────────────────────────────────────────────────────────────
# #  STEP 3 – PREVIEW & DOWNLOAD
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="step-card">
#   <div class="step-header">
#     <span class="step-num">3</span>
#     <span class="step-title">Preview &amp; Download</span>
#   </div>
# </div>
# """, unsafe_allow_html=True)

# # ── Re-trim using the user's actual mapped columns (catches any edge cases) ──
# df = trim_data_rows(df, mapping.get("stock_no"), mapping.get("carat"))

# records = []
# for _, row in df.iterrows():
#     rec = {key: (row[col] if col else None) for key, col in mapping.items()}
#     records.append(rec)

# mapped_df = pd.DataFrame(records)
# for col in ["carat", "final", "vdb", "usa"]:
#     mapped_df[col] = pd.to_numeric(mapped_df[col], errors="coerce")

# mapped_df = mapped_df.dropna(subset=["carat", "final", "vdb", "usa"], how="all")
# n_stones  = len(mapped_df)

# if n_stones == 0:
#     st.markdown('<div class="banner banner-warn">⚠️ No valid rows found. Check your column mappings.</div>',
#                 unsafe_allow_html=True)
#     st.stop()

# st.markdown(f'<div class="banner banner-success">✅ <b>{n_stones} stones</b> ready to export.</div>',
#             unsafe_allow_html=True)

# preview_rename = {
#     "stock_no":"Stock ID.", "certi_no":"Certi No.", "shape":"Shape",
#     "carat":"Carat", "lab":"Lab", "color":"Color", "clarity":"Clarity",
#     "final":"Final", "vdb":"VDB", "usa":"USA",
# }
# st.dataframe(mapped_df.rename(columns=preview_rename).head(10), use_container_width=True)

# with st.spinner("Building Excel with formulas and colours…"):
#     xlsx_bytes   = build_output_excel(mapped_df, rate_value=rate_input, tariff_value=tariff_input)
#     out_filename = safe_filename(party_file.name, n_stones)

# st.markdown('<div class="banner banner-success">💾 File ready — all formulas match the reference format exactly (cols P–Y yellow).</div>',
#             unsafe_allow_html=True)

# st.download_button(
#     label=f"⬇️  Download  {out_filename}",
#     data=xlsx_bytes,
#     file_name=out_filename,
#     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#     use_container_width=True,
# )

# with st.expander("📐 Formula & Colour Reference"):
#     st.markdown(f"""
# | Col | Header | Colour | Formula / Source |
# |---|---|---|---|
# | A | Stock ID | — | From party file |
# | B | Certi No | — | From party file |
# | C | Shape | — | From party file |
# | D | Carat | 🟦 Teal | From party file |
# | E | Lab | — | From party file |
# | F | Color | — | From party file |
# | G | Clarity | — | From party file |
# | H | Final | 🟨 Gold | From party file |
# | I | Amt | 🟨 Gold | `D × H` |
# | J | VDB | 🔵 Blue | From party file |
# | K | Amt | 🔵 Blue | `J × D` |
# | L | USA | 🟠 Orange | From party file |
# | M | Amt | 🟠 Orange | `L × D` |
# | N | With Tariff | 🟣 Lavender | `H × Tariff Multiplier` — Multiplier = **{tariff_input}** |
# | O | Amt | 🟣 Lavender | `N × D` |
# | P | — | 🟡 Yellow | `ROUND(((L-(M_avg−O_avg))/1.08),2)` |
# | Q | — | 🟡 Yellow | `ROUND(P×D, 2)` |
# | R | — | 🟡 Yellow | `ROUND(P×1.08, 2)` |
# | S | — | 🟡 Yellow | `ROUND(L−R, 2)` |
# | T | — | 🟡 Yellow | `ROUND((S×100)/P, 2)` |
# | U | — | 🟡 Yellow | `ROUND((L×ΣI) / SUMPRODUCT(ΣL,ΣD), 2)` |
# | V | — | 🟡 Yellow | `ROUND(U×D, 2)` |
# | W | — | 🟡 Yellow | `ROUND(L−U, 2)` |
# | X | — | 🟡 Yellow | `ROUND((W×100)/U, 2)` |
# | Y | — | 🟡 Yellow | `ROUND(U×$RATE, 2)` — RATE = **{rate_input}** |

# > 🟡 Both **TARIFF x** and **$ RATE** cells are highlighted yellow/red in the summary block — editable directly in Excel.
# """)


# import streamlit as st
# import pandas as pd
# import openpyxl
# from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
# from openpyxl.utils import get_column_letter
# import io
# import re
# import os
# from datetime import datetime

# # ─────────────────────────────────────────────────────────────────────────────
# #  PAGE CONFIG
# # ─────────────────────────────────────────────────────────────────────────────
# st.set_page_config(
#     page_title="💎 Diamond Formatter",
#     page_icon="💎",
#     layout="wide",
#     initial_sidebar_state="collapsed",
# )

# # ─────────────────────────────────────────────────────────────────────────────
# #  DARK THEME CSS
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <style>
# html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"] {
#     background-color: #0d0f14 !important; color: #e2e8f0 !important;
# }
# [data-testid="stHeader"] { background: #0d0f14 !important; }
# [data-testid="stSidebar"] { background: #111318 !important; }
# section[data-testid="stMain"] > div { background: #0d0f14 !important; }
# .block-container { padding-top: 1.5rem !important; }
# .hero {
#     position: relative;
#     background: linear-gradient(135deg, #0f1629 0%, #1a0a2e 50%, #0a1628 100%);
#     background-size: 200% 200%;
#     animation: heroGradientShift 12s ease infinite;
#     border: 1px solid #2a2d3e; border-radius: 20px;
#     padding: 2.4rem 2rem; margin-bottom: 2rem; overflow: hidden;
#     text-align: center; box-shadow: 0 10px 40px rgba(0,0,0,.35);
# }
# @keyframes heroGradientShift {
#     0%   { background-position: 0% 50%; }
#     50%  { background-position: 100% 50%; }
#     100% { background-position: 0% 50%; }
# }
# .hero::before, .hero::after {
#     content: ''; position: absolute; border-radius: 50%;
#     filter: blur(60px); opacity: .55;
#     animation: heroFloat 8s ease-in-out infinite; pointer-events: none;
# }
# .hero::before { width:240px;height:240px;background:rgba(99,102,241,.35);top:-70px;left:8%; }
# .hero::after  { width:280px;height:280px;background:rgba(168,85,247,.28);bottom:-90px;right:8%;animation-delay:-4s; }
# @keyframes heroFloat {
#     0%,100% { transform:translateY(0) translateX(0); }
#     50%     { transform:translateY(-22px) translateX(16px); }
# }
# .hero-content { position:relative;z-index:1;text-align:left; }
# .hero-company {
#     display:inline-block;font-size:.92rem;font-weight:700;letter-spacing:2.6px;
#     color:#c4b5fd;text-transform:uppercase;padding:7px 20px;
#     border:1px solid rgba(167,139,250,.45);border-radius:30px;
#     background:rgba(99,102,241,.12);margin:0 0 1.1rem;
#     animation:heroFadeInDown .8s ease both,heroPulseBorder 3s ease-in-out infinite 1s;
# }
# @keyframes heroFadeInDown { from{opacity:0;transform:translateY(-14px)} to{opacity:1;transform:translateY(0)} }
# @keyframes heroPulseBorder {
#     0%,100% { box-shadow:0 0 0 rgba(167,139,250,0); }
#     50%     { box-shadow:0 0 20px rgba(167,139,250,.4); }
# }
# .hero-title { display:flex;align-items:center;justify-content:flex-start;gap:12px;margin:0 0 .6rem;animation:heroFadeInUp .9s ease .2s both; }
# .hero-gem-wrap { position:relative;display:inline-flex;align-items:center;justify-content:center; }
# .hero-title .gem { display:inline-block;font-size:2rem;line-height:1;transform-style:preserve-3d;animation:heroGemSpin 4.5s ease-in-out infinite; }
# .hero-gem-wrap .spark { position:absolute;color:#7dd3fc;font-size:.6rem;opacity:0; }
# .hero-gem-wrap .spark-1 { top:-6px;right:-8px;animation:heroTwinkle 2.6s ease-in-out infinite; }
# .hero-gem-wrap .spark-2 { bottom:-4px;left:-10px;animation:heroTwinkle 2.6s ease-in-out infinite 1.1s; }
# @keyframes heroTwinkle { 0%,100%{opacity:0;transform:scale(.4) rotate(0deg)} 50%{opacity:1;transform:scale(1.1) rotate(25deg)} }
# @keyframes heroGemSpin {
#     0%   { transform:perspective(320px) rotateY(0deg) scale(1);filter:drop-shadow(0 0 6px rgba(167,139,250,.45)); }
#     45%  { transform:perspective(320px) rotateY(180deg) scale(1.1);filter:drop-shadow(0 0 16px rgba(56,189,248,.8)); }
#     55%  { transform:perspective(320px) rotateY(180deg) scale(1.1);filter:drop-shadow(0 0 16px rgba(56,189,248,.8)); }
#     100% { transform:perspective(320px) rotateY(360deg) scale(1);filter:drop-shadow(0 0 6px rgba(167,139,250,.45)); }
# }
# .hero-title .txt {
#     font-size:2.3rem;font-weight:800;letter-spacing:-.5px;line-height:1.15;
#     background:linear-gradient(90deg,#a78bfa,#6366f1,#38bdf8,#a78bfa);
#     background-size:300% auto;
#     -webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;
#     animation:heroShimmer 6s linear infinite;
# }
# @keyframes heroShimmer { to{background-position:300% center} }
# @keyframes heroFadeInUp { from{opacity:0;transform:translateY(18px)} to{opacity:1;transform:translateY(0)} }
# .hero-sub { color:#94a3b8;font-size:1rem;margin:0;animation:heroFadeInUp .9s ease .45s both; }
# .legend-wrap { display:flex;flex-wrap:wrap;gap:8px;margin:1rem 0 1.5rem; }
# .pill {
#     display:inline-flex;align-items:center;gap:7px;
#     background:#1a1d27;border:1px solid #2a2d3e;
#     border-radius:20px;padding:5px 12px;font-size:.78rem;color:#cbd5e1;
#     transition:transform .18s ease,border-color .18s ease;
# }
# .pill:hover { transform:translateY(-2px);border-color:#6366f1; }
# .dot { width:12px;height:12px;border-radius:3px;flex-shrink:0; }
# .step-card { background:#13161f;border:1px solid #1e2130;border-radius:12px;padding:1.5rem;margin-bottom:1rem; }
# .step-num {
#     display:inline-flex;align-items:center;justify-content:center;
#     width:32px;height:32px;border-radius:50%;
#     background:linear-gradient(135deg,#6366f1,#8b5cf6);
#     color:white;font-weight:700;font-size:.9rem;margin-right:10px;flex-shrink:0;
# }
# .step-title { font-size:1.1rem;font-weight:600;color:#f1f5f9; }
# .step-header { display:flex;align-items:center;margin-bottom:1rem; }
# [data-testid="stFileUploader"] { background:#13161f !important;border:2px dashed #2a2d3e !important;border-radius:10px !important; }
# [data-testid="stFileUploader"]:hover { border-color:#6366f1 !important; }
# .stSelectbox > div > div { background:#1a1d27 !important;border:1px solid #2a2d3e !important;color:#e2e8f0 !important;border-radius:8px !important; }
# .stSelectbox label { color:#cbd5e1 !important;font-size:.85rem !important; }
# .stDownloadButton button {
#     background:linear-gradient(135deg,#6366f1,#8b5cf6) !important;
#     color:white !important;border:none !important;border-radius:10px !important;
#     font-weight:600 !important;padding:.75rem 1.5rem !important;
#     font-size:1rem !important;transition:all .2s !important;
#     box-shadow:0 4px 20px rgba(99,102,241,.4) !important;
# }
# .stDownloadButton button:hover { transform:translateY(-2px) !important;box-shadow:0 8px 30px rgba(99,102,241,.6) !important; }
# .banner { border-radius:10px;padding:.85rem 1.1rem;font-size:.9rem;margin:.5rem 0;display:flex;align-items:center;gap:10px; }
# .banner-success { background:rgba(34,197,94,.1);border:1px solid rgba(34,197,94,.3);color:#86efac; }
# .banner-info    { background:rgba(99,102,241,.1);border:1px solid rgba(99,102,241,.3);color:#a5b4fc; }
# .banner-warn    { background:rgba(234,179,8,.1); border:1px solid rgba(234,179,8,.3); color:#fde68a; }
# .stats-row { display:flex;gap:12px;flex-wrap:wrap;margin:1rem 0; }
# .stat-box { background:#13161f;border:1px solid #1e2130;border-radius:10px;padding:.75rem 1.25rem;flex:1;min-width:130px; }
# .stat-val { font-size:1.6rem;font-weight:700;color:#a78bfa; }
# .stat-label { font-size:.75rem;color:#64748b;text-transform:uppercase;letter-spacing:.5px; }
# hr { border-color:#1e2130 !important;margin:1.5rem 0 !important; }
# [data-testid="stExpander"] { background:#13161f !important;border:1px solid #1e2130 !important;border-radius:10px !important; }
# summary { color:#94a3b8 !important; }
# </style>
# """, unsafe_allow_html=True)

# # ─────────────────────────────────────────────────────────────────────────────
# #  COLOR CONSTANTS
# #  Matches Image 2 exactly:
# #  - Cols A-O : existing group colors (navy header, teal, gold, blue, orange, lavender)
# #  - Cols P-Y : YELLOW (FFFF00) — boss requirement
# # ─────────────────────────────────────────────────────────────────────────────
# CLR_HEADER  = "1F4E79"   # deep navy header
# CLR_CARAT   = "98F6EB"   # teal   → col D (4)
# CLR_FINAL   = "FFCF37"   # gold   → H,I (8,9)
# CLR_VDB     = "5991D5"   # blue   → J,K (10,11)
# CLR_USA     = "F79B4F"   # orange → L,M (12,13)
# CLR_TARIFF  = "9E89B9"   # lavender → N,O (14,15)
# CLR_YELLOW  = "FFFF00"   # yellow → P-Y (16-25)

# # col index (1-based) → fill color for DATA cells
# COL_COLORS = {
#     4:  CLR_CARAT,
#     8:  CLR_FINAL,   9:  CLR_FINAL,
#     10: CLR_VDB,     11: CLR_VDB,
#     12: CLR_USA,     13: CLR_USA,
#     14: CLR_TARIFF,  15: CLR_TARIFF,
# }
# # cols 16-26 all yellow
# for _c in range(16, 27):
#     COL_COLORS[_c] = CLR_YELLOW

# # ─────────────────────────────────────────────────────────────────────────────
# #  OUTPUT HEADERS  (25 columns, matching Image 2)
# # ─────────────────────────────────────────────────────────────────────────────
# OUTPUT_HEADERS = [
#     "Stock ID.",        # A  1
#     "Certi No.",        # B  2
#     "Shape",            # C  3
#     "Carat",            # D  4
#     "Lab",              # E  5
#     "Color",            # F  6
#     "Clarity",          # G  7
#     "Final",            # H  8
#     "Amt",              # I  9
#     "VDB",              # J  10
#     "Amt",              # K  11
#     "USA",              # L  12
#     "Amt",              # M  13
#     "With Tariff",      # N  14
#     "Amt",              # O  15
#     "",                 # P  16  (yellow — unlabeled in source)
#     "",                 # Q  17
#     "",                 # R  18
#     "",                 # S  19
#     "",                 # T  20
#     "",                 # U  21
#     "",                 # V  22
#     "",                 # W  23
#     "",                 # X  24
#     "",                 # Y  25
#     "Total INR Amt",   # Z  26
# ]

# # NOTE: "with_tariff" is no longer taken from the party file — it is now
# # CALCULATED as Final × Tariff-Multiplier (user enters the multiplier, e.g.
# # 1.08 or 1.06), so it has been removed from the columns the user maps.
# REQUIRED_COLS = [
#     ("stock_no",    "Stock ID."),
#     ("certi_no",    "Certi No."),
#     ("shape",       "Shape"),
#     ("carat",       "Carat"),
#     ("lab",         "Lab"),
#     ("color",       "Color"),
#     ("clarity",     "Clarity"),
#     ("final",       "Final"),
#     ("vdb",         "VDB"),
#     ("usa",         "USA"),
# ]

# # ─────────────────────────────────────────────────────────────────────────────
# #  HELPERS
# # ─────────────────────────────────────────────────────────────────────────────
# def read_uploaded_file(f) -> pd.DataFrame | None:
#     name = f.name.lower()
#     try:
#         if name.endswith(".csv"):
#             return pd.read_csv(f)
#         elif name.endswith((".xlsx", ".xls", ".xlsm")):
#             return pd.read_excel(f, header=None)
#         elif name.endswith(".tsv"):
#             return pd.read_csv(f, sep="\t")
#         else:
#             st.error("Unsupported format. Please upload CSV, XLSX, XLS, XLSM, or TSV.")
#             return None
#     except Exception as e:
#         st.error(f"Could not read file: {e}")
#         return None


# def detect_header_row(df: pd.DataFrame) -> int:
#     for i, row in df.iterrows():
#         vals = [str(v).strip() for v in row if pd.notna(v) and str(v).strip()]
#         if len(vals) >= 4 and not all(
#             re.sub(r'[.\-]', '', v).isdigit() for v in vals
#         ):
#             return i
#     return 0


# def trim_data_rows(df: pd.DataFrame, stock_col: str | None, carat_col: str | None) -> pd.DataFrame:
#     """
#     Drop any trailing non-data rows (SUM totals, labels, blanks) that appear
#     after the real diamond records.

#     A row is treated as a SUM / footer row — and all rows from it onward are
#     dropped — when ANY of these conditions holds:

#     1. Carat column is mapped AND the carat cell has a numeric value BUT the
#        stock-ID cell is empty or NaN.  This is the classic pattern:
#            Stock  | Carat
#            (blank)| 28.41   ← SUM row
#     2. Carat column is mapped AND the carat value is ≥ 3× the median carat of
#        the rows seen so far (catches a huge SUM value even if stock is filled).
#     3. Stock-ID column is mapped AND the stock-ID cell is a purely numeric string
#        or a known summary keyword (TOTAL, AVG, SUM, RATE, FINAL, etc.).
#     """
#     if df.empty:
#         return df

#     summary_keywords = {"total", "avg", "average", "sum", "rate", "final",
#                         "difference", "subtotal", "done", "fianl"}

#     # Compute median carat from first 5 valid rows for outlier detection
#     median_carat = None
#     if carat_col and carat_col in df.columns:
#         sample_carats = pd.to_numeric(df[carat_col].head(20), errors="coerce").dropna()
#         if not sample_carats.empty:
#             median_carat = sample_carats.median()

#     cut_at = len(df)  # default: keep everything

#     for i, (_, row) in enumerate(df.iterrows()):
#         # ── condition 1: empty stock + non-empty carat ──
#         if stock_col and stock_col in df.columns and carat_col and carat_col in df.columns:
#             stock_val = row.get(stock_col)
#             carat_val = row.get(carat_col)
#             stock_empty = pd.isna(stock_val) or str(stock_val).strip() == ""
#             carat_numeric = pd.notna(carat_val) and str(carat_val).strip() != ""
#             if stock_empty and carat_numeric:
#                 cut_at = i
#                 break

#         # ── condition 2: carat value is a huge outlier (SUM) ──
#         if carat_col and carat_col in df.columns and median_carat:
#             carat_val = pd.to_numeric(row.get(carat_col), errors="coerce")
#             if pd.notna(carat_val) and carat_val >= median_carat * 3:
#                 cut_at = i
#                 break

#         # ── condition 3: stock ID is a summary keyword ──
#         if stock_col and stock_col in df.columns:
#             stock_val = str(row.get(stock_col, "")).strip().lower()
#             if stock_val in summary_keywords:
#                 cut_at = i
#                 break

#     return df.iloc[:cut_at].reset_index(drop=True)


# def safe_filename(raw_name: str, n_stones: int) -> str:
#     """Party_Name_No_of_Stones_Current_Date.xlsx"""
#     stem = os.path.splitext(raw_name)[0]
#     stem = re.sub(r'[^\w\s\-]', '', stem).strip()
#     stem = re.sub(r'\s+', '_', stem)
#     date_str = datetime.now().strftime("%d-%m-%Y")
#     return f"{stem}_{n_stones}_Stones_{date_str}.xlsx"


# def smart_default(key: str, cols: list) -> int:
#     hints = {
#         "stock_no":    ["stock", "pexkt", "packet", "pkt", "id", "lot"],
#         "certi_no":    ["certi", "cert", "gia", "igi", "report"],
#         "shape":       ["shape"],
#         "carat":       ["carat", "ct", "weight"],
#         "lab":         ["lab", "igi", "gia"],
#         "color":       ["color", "colour"],
#         "clarity":     ["clarity", "clar"],
#         "final":       ["final", "rap%", "disc%", "back%"],
#         "vdb":         ["vdb"],
#         "usa":         ["usa"],
#     }
#     keywords = hints.get(key, [])
#     lower_cols = [c.lower() for c in cols[1:]]
#     for kw in keywords:
#         for i, lc in enumerate(lower_cols):
#             if kw in lc:
#                 return i + 1
#     return 0


# def px(hex_color: str) -> PatternFill:
#     return PatternFill("solid", fgColor=hex_color)


# def thin_border() -> Border:
#     s = Side(style="thin", color="D0D0D0")
#     return Border(left=s, right=s, top=s, bottom=s)


# def add_calculated_fields(df: pd.DataFrame, tariff_value: float, rate_value: float) -> pd.DataFrame:
#     """
#     Mirrors the exact Excel formulas (N, O, P..Z) in pandas so we can:
#       1) sort the data rows smallest → highest by column U before writing
#          them into the sheet (same order the U formula would produce), and
#       2) build the Summary sheet's insights/charts from the same numbers.
#     This does NOT change any Excel formula written by build_output_excel —
#     it's a read-only mirror used purely for ordering + reporting.
#     """
#     d = df.copy()
#     d["final_amt"]  = d["carat"] * d["final"]                 # I
#     d["vdb_amt"]    = d["vdb"] * d["carat"]                   # K
#     d["usa_amt"]    = d["usa"] * d["carat"]                   # M
#     d["tariff"]     = d["final"] * tariff_value                # N
#     d["tariff_amt"] = d["tariff"] * d["carat"]                 # O

#     sum_final_amt = d["final_amt"].sum()
#     sumproduct_ld = (d["usa"] * d["carat"]).sum()

#     d["U"] = (d["usa"] * sum_final_amt / sumproduct_ld) if sumproduct_ld else 0.0
#     d["V"] = d["U"] * d["carat"]
#     d["W"] = d["usa"] - d["U"]
#     d["X"] = (d["W"] * 100) / d["U"].replace(0, pd.NA)
#     d["Y"] = d["U"] * rate_value
#     d["total_inr"] = d["Y"] * d["carat"]                       # Z — Total INR Amt
#     return d


# # ─────────────────────────────────────────────────────────────────────────────
# #  CORE EXCEL BUILDER
# #
# #  Layout (matches Image 2 / source file exactly):
# #   Row 1        : Headers
# #   Rows 2..DS+n-1 : Data  (n = number of stones)
# #   SUM_ROW      : TOTAL
# #   AVG_ROW      : AVG / CARAT
# #   (blank rows)
# #   FINAL_ROW    : FINAL label + value
# #   TARIFF_ROW   : TARIFF x  label + value (yellow, red text) ← used by col-N formula
# #   RATE_ROW     : $ RATE label + value (yellow, red text) ← used by col-Y formula
# #   DIFF_ROW     : Difference
# #
# #  Column formulas (H onward) replicate source exactly, adjusted for output row refs:
# #   N  =ROUND(H*$N$TARIFF_ROW, 2)                 ← With Tariff = Final × Tariff multiplier
# #   O  =ROUND(N*D, 2)
# #   P  =ROUND(((L-(M_AVG - O_AVG))/1.08), 2)      ← where M_AVG=M$AVG_ROW, O_AVG=O$AVG_ROW
# #   Q  =ROUND(P*D, 2)
# #   R  =ROUND(P*1.08, 2)
# #   S  =ROUND(L-R, 2)
# #   T  =ROUND((S*100)/P, 2)
# #   U  =ROUND((L*$I$SUM_ROW)/SUMPRODUCT($L$DS:$L$DE,$D$DS:$D$DE), 2)
# #   V  =ROUND(U*D, 2)
# #   W  =ROUND(L-U, 2)
# #   X  =ROUND((W*100)/U, 2)
# #   Y  =ROUND(U*$N$RATE_ROW, 2)   ← references RATE cell instead of hardcoded 95.95
# #
# #  SUM row  : D, I, K, M, O, Q, V
# #  AVG row  : I, K, M, O, Q, V  (each SUM/D_SUM)
# # ─────────────────────────────────────────────────────────────────────────────
# def build_output_excel(mapped_df: pd.DataFrame, rate_value: float = 95.95,
#                         tariff_value: float = 1.08) -> bytes:
#     wb = openpyxl.Workbook()
#     ws = wb.active
#     ws.title = "Diamond Data"

#     n        = len(mapped_df)
#     HDR      = 1
#     DS       = 2            # data start row
#     DE       = DS + n - 1   # data end row
#     SUM_ROW    = DE + 1
#     AVG_ROW    = DE + 2
#     FINAL_ROW  = DE + 5
#     TARIFF_ROW = DE + 6
#     RATE_ROW   = DE + 7
#     DIFF_ROW   = DE + 8

#     # ── Header row ────────────────────────────────────────────────────────
#     for c, hdr in enumerate(OUTPUT_HEADERS, 1):
#         cell = ws.cell(row=HDR, column=c, value=hdr)
#         cell.fill      = px(CLR_HEADER)
#         cell.font      = Font(name="Arial", bold=True, color="FFFFFF", size=10)
#         cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
#         cell.border    = thin_border()
#     ws.row_dimensions[HDR].height = 36

#     # ── Data rows ─────────────────────────────────────────────────────────
#     for i, (_, row) in enumerate(mapped_df.iterrows()):
#         r = DS + i

#         # Raw input columns (A-M in output = cols 1-13, from party file)
#         raw = {
#             1:  row.get("stock_no"),
#             2:  row.get("certi_no"),
#             3:  row.get("shape"),
#             4:  row.get("carat"),
#             5:  row.get("lab"),
#             6:  row.get("color"),
#             7:  row.get("clarity"),
#             8:  row.get("final"),       # H  Final (%)
#             10: row.get("vdb"),         # J  VDB (%)
#             12: row.get("usa"),         # L  USA (%)
#         }

#         # Formula columns — col N (14) is now CALCULATED as Final × Tariff
#         # multiplier instead of being read from the party file.
#         formulas = {
#             9:  f"=ROUND(D{r}*H{r},2)",           # I  Final Amt
#             11: f"=ROUND(J{r}*D{r},2)",            # K  VDB Amt
#             13: f"=ROUND(L{r}*D{r},2)",            # M  USA Amt

#             # N  With Tariff = Final × Tariff multiplier (user-entered, cell ref)
#             14: f"=ROUND(H{r}*$N${TARIFF_ROW},2)",
#             # O  Tariff Amt = N*D
#             15: f"=ROUND(N{r}*D{r},2)",

#             # ── YELLOW columns (P-Y) — exact source formulas ──
#             # P: =ROUND(((L-(M_AVG-O_AVG))/1.08),2)
#             16: f"=ROUND(((L{r}-(M${AVG_ROW}-O${AVG_ROW}))/1.08),2)",
#             # Q: =ROUND(P*D,2)
#             17: f"=ROUND(P{r}*D{r},2)",
#             # R: =ROUND(P*1.08,2)
#             18: f"=ROUND(P{r}*1.08,2)",
#             # S: =ROUND(L-R,2)
#             19: f"=ROUND(L{r}-R{r},2)",
#             # T: =ROUND((S*100)/P,2)
#             20: f"=ROUND((S{r}*100)/P{r},2)",
#             # U: =ROUND((L*$I$SUM)/SUMPRODUCT($L$DS:$L$DE,$D$DS:$D$DE),2)
#             21: f"=ROUND((L{r}*$I${SUM_ROW})/SUMPRODUCT($L${DS}:$L${DE},$D${DS}:$D${DE}),2)",
#             # V: =ROUND(U*D,2)
#             22: f"=ROUND(U{r}*D{r},2)",
#             # W: =ROUND(L-U,2)
#             23: f"=ROUND(L{r}-U{r},2)",
#             # X: =ROUND((W*100)/U,2)
#             24: f"=ROUND((W{r}*100)/U{r},2)",
#             # Y: =ROUND(U*RATE,2)  — references the RATE cell
#             25: f"=ROUND(U{r}*$N${RATE_ROW},2)",
#             # Z: =ROUND(Y*D,2)  — Total INR Amt
#             26: f"=ROUND(Y{r}*D{r},2)",
#         }

#         for c in range(1, 27):
#             val = raw.get(c, formulas.get(c))
#             cell = ws.cell(row=r, column=c, value=val)
#             cell.border    = thin_border()
#             cell.alignment = Alignment(horizontal="center", vertical="center")
#             cell.font      = Font(name="Arial", size=10)
#             if c in COL_COLORS:
#                 cell.fill = px(COL_COLORS[c])
#             if c in (4, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26):
#                 cell.number_format = "0.00"

#         ws.row_dimensions[r].height = 18

#     # ── TOTAL row ─────────────────────────────────────────────────────────
#     total_fill = px("F2F2F2")
#     lbl = ws.cell(row=SUM_ROW, column=1, value="TOTAL")
#     lbl.font = Font(name="Arial", bold=True, size=10)
#     lbl.fill = total_fill; lbl.border = thin_border()
#     lbl.alignment = Alignment(horizontal="center")

#     # SUM cols: D(4), I(9), K(11), M(13), O(15), Q(17), V(22), Z(26)
#     sum_cols = {4:"D", 9:"I", 11:"K", 13:"M", 15:"O", 17:"Q", 22:"V", 26:"Z"}
#     for col, letter in sum_cols.items():
#         cell = ws.cell(row=SUM_ROW, column=col,
#                        value=f"=SUM({letter}{DS}:{letter}{DE})")
#         cell.font = Font(name="Arial", bold=True, size=10)
#         cell.fill = total_fill; cell.border = thin_border()
#         cell.alignment = Alignment(horizontal="center")
#         cell.number_format = "0.00"

#     # ── AVG / CARAT row ───────────────────────────────────────────────────
#     avg_fill = px("E2EFDA")
#     lbl2 = ws.cell(row=AVG_ROW, column=1, value="AVG / CARAT")
#     lbl2.font = Font(name="Arial", bold=True, size=10)
#     lbl2.fill = avg_fill; lbl2.border = thin_border()
#     lbl2.alignment = Alignment(horizontal="center")

#     # AVG cols: I(9), K(11), M(13), O(15), Q(17), V(22)
#     avg_cols = {9:"I", 11:"K", 13:"M", 15:"O", 17:"Q", 22:"V"}
#     for col, letter in avg_cols.items():
#         cell = ws.cell(row=AVG_ROW, column=col,
#                        value=f"=ROUND({letter}{SUM_ROW}/D{SUM_ROW},2)")
#         cell.font = Font(name="Arial", bold=True, size=10)
#         cell.fill = avg_fill; cell.border = thin_border()
#         cell.alignment = Alignment(horizontal="center")
#         cell.number_format = "0.00"

#     # ── FINAL / TARIFF / RATE / Difference label block ───────────────────
#     # FINAL row
#     for c, val in ((13, "FINAL"), (14, f"=ROUND(I{AVG_ROW},2)")):
#         cell = ws.cell(row=FINAL_ROW, column=c, value=val)
#         cell.font = Font(name="Arial", bold=True, size=10)
#         cell.border = thin_border()
#         cell.alignment = Alignment(horizontal="center")
#         cell.number_format = "0.00"

#     # TARIFF row — label M, value N (yellow fill, red bold text for value)
#     # This is the multiplier used to CALCULATE "With Tariff" (col N) as
#     # Final × Tariff, e.g. 1.08 or 1.06 — editable directly in Excel.
#     lbl_tariff = ws.cell(row=TARIFF_ROW, column=13, value="TARIFF x")
#     lbl_tariff.font = Font(name="Arial", bold=True, size=10)
#     lbl_tariff.border = thin_border()
#     lbl_tariff.alignment = Alignment(horizontal="center")

#     tariff_cell = ws.cell(row=TARIFF_ROW, column=14, value=tariff_value)
#     tariff_cell.fill   = px("FFFF00")
#     tariff_cell.font   = Font(name="Arial", bold=True, color="FF0000", size=10)
#     tariff_cell.border = thin_border()
#     tariff_cell.alignment = Alignment(horizontal="center")
#     tariff_cell.number_format = "0.00"

#     # RATE row — label M, value N (yellow fill, red bold text for value)
#     lbl_rate = ws.cell(row=RATE_ROW, column=13, value="$ RATE")
#     lbl_rate.font = Font(name="Arial", bold=True, size=10)
#     lbl_rate.border = thin_border()
#     lbl_rate.alignment = Alignment(horizontal="center")

#     rate_cell = ws.cell(row=RATE_ROW, column=14, value=rate_value)
#     rate_cell.fill   = px("FFFF00")
#     rate_cell.font   = Font(name="Arial", bold=True, color="FF0000", size=10)
#     rate_cell.border = thin_border()
#     rate_cell.alignment = Alignment(horizontal="center")
#     rate_cell.number_format = "0.00"

#     # Difference row
#     for c, val in ((13, "Difference"), (14, f"=M{AVG_ROW}-O{AVG_ROW}")):
#         cell = ws.cell(row=DIFF_ROW, column=c, value=val)
#         cell.font = Font(name="Arial", bold=True, size=10)
#         cell.border = thin_border()
#         cell.alignment = Alignment(horizontal="center")
#         cell.number_format = "0.00"

#     # ── Column widths ──────────────────────────────────────────────────────
#     widths = [14, 14, 13, 7, 6, 7, 9,  8, 11,  8, 11,  8, 11, 12, 11,
#               11, 11, 11,  8,  9, 11, 11, 11,  9, 13, 15]
#     for i, w in enumerate(widths, 1):
#         ws.column_dimensions[get_column_letter(i)].width = w

#     ws.freeze_panes = "A2"

#     # ── AutoFilter across the header + data table (A:Z) ─────────────────────
#     ws.auto_filter.ref = f"A{HDR}:Z{DE}"

#     buf = io.BytesIO()
#     wb.save(buf)
#     buf.seek(0)
#     return buf.getvalue()


# # ─────────────────────────────────────────────────────────────────────────────
# #  HERO BANNER
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="hero">
#   <div class="hero-content">
#     <p class="hero-company">✦ BRIGHT TRADING INDIA LLP</p>
#     <p class="hero-title">
#       <span class="hero-gem-wrap">
#         <span class="gem">💎</span>
#         <span class="spark spark-1">✦</span>
#         <span class="spark spark-2">✦</span>
#       </span>
#       <span class="txt">Diamond Cost Optimizer</span>
#     </p>
#     <p class="hero-sub">Smart Cost Adjustment &amp; Equal Profit Calculator</p>
#   </div>
# </div>
# """, unsafe_allow_html=True)

# st.markdown("""
# <div class="legend-wrap">
#   <div class="pill"><div class="dot" style="background:#98F6EB"></div>Carat</div>
#   <div class="pill"><div class="dot" style="background:#FFCF37"></div>Final / Amt</div>
#   <div class="pill"><div class="dot" style="background:#5991D5"></div>VDB / Amt</div>
#   <div class="pill"><div class="dot" style="background:#F79B4F"></div>USA / Amt</div>
#   <div class="pill"><div class="dot" style="background:#9E89B9"></div>With Tariff / Amt</div>
#   <div class="pill"><div class="dot" style="background:#FFFF00"></div>Calculated Columns (P–Y)</div>
# </div>
# """, unsafe_allow_html=True)

# # ─────────────────────────────────────────────────────────────────────────────
# #  STEP 1 – UPLOAD
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="step-card">
#   <div class="step-header">
#     <span class="step-num">1</span>
#     <span class="step-title">Upload Party File</span>
#   </div>
#   <p style="color:#64748b;font-size:.87rem;margin:0 0 .9rem 42px;">
#     Accepts <b style="color:#a78bfa">CSV, XLSX, XLS, XLSM, TSV</b> — any column order, any layout.
#   </p>
# </div>
# """, unsafe_allow_html=True)

# party_file = st.file_uploader(
#     "Drop your file here or click Upload",
#     type=["csv", "xlsx", "xls", "xlsm", "tsv"],
#     key="party_file",
#     label_visibility="collapsed",
# )

# if party_file is None:
#     st.markdown('<div class="banner banner-info">👆 Upload the party file above to begin.</div>',
#                 unsafe_allow_html=True)
#     st.stop()

# raw_df = read_uploaded_file(party_file)
# if raw_df is None:
#     st.stop()

# h_row = detect_header_row(raw_df)
# df = raw_df.copy()
# df.columns = raw_df.iloc[h_row]
# df = df.iloc[h_row + 1:].reset_index(drop=True)
# df.columns = [str(c).strip() if pd.notna(c) else f"Col_{i}" for i, c in enumerate(df.columns)]
# df = df.dropna(how="all")

# # ── Auto-detect and strip SUM/footer rows at the bottom ──────────────────────
# # We do a quick keyword-based column sniff so trim works before Step-2 mapping
# def _sniff_col(df, keywords):
#     for col in df.columns:
#         if any(kw in str(col).lower() for kw in keywords):
#             return col
#     return None

# _stock_col = _sniff_col(df, ["stock", "packet", "pkt", "lot", "id"])
# _carat_col = _sniff_col(df, ["carat", "ct", "weight"])
# df = trim_data_rows(df, _stock_col, _carat_col)

# st.markdown(f"""
# <div class="stats-row">
#   <div class="stat-box"><div class="stat-val">{len(df)}</div><div class="stat-label">Rows Detected</div></div>
#   <div class="stat-box"><div class="stat-val">{len(df.columns)}</div><div class="stat-label">Columns Found</div></div>
#   <div class="stat-box"><div class="stat-val">{party_file.name.rsplit(".",1)[-1].upper()}</div><div class="stat-label">File Format</div></div>
# </div>
# """, unsafe_allow_html=True)

# with st.expander("👀 Preview raw data (first 5 rows)"):
#     st.dataframe(df.head(5), use_container_width=True)

# # ─────────────────────────────────────────────────────────────────────────────
# #  STEP 2 – COLUMN MAPPING
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="step-card">
#   <div class="step-header">
#     <span class="step-num">2</span>
#     <span class="step-title">Map Columns</span>
#   </div>
# </div>
# """, unsafe_allow_html=True)
# st.caption("Select which column in the party file corresponds to each required field.")

# available = ["— skip / not available —"] + list(df.columns)
# mapping   = {}

# cols_left, cols_right = st.columns(2)
# for idx, (key, label) in enumerate(REQUIRED_COLS):
#     default = smart_default(key, available)
#     container = cols_left if idx % 2 == 0 else cols_right
#     with container:
#         choice = st.selectbox(f"**{label}**", options=available,
#                               index=default, key=f"map_{key}")
#         mapping[key] = None if choice.startswith("—") else choice

# # ── Tariff multiplier + $ RATE inputs ─────────────────────────────────────────
# st.markdown("---")
# # st.markdown("**Tariff Multiplier** — used to calculate the `With Tariff` column (`Final × Multiplier`), e.g. 1.08 or 1.06")
# tariff_input = st.number_input("Tariff multiplier", min_value=1.0, value=1.08, step=0.01, format="%.2f")

# # st.markdown("**$ Rate** — used in the final Y-column formula (`U × $ Rate`)")
# rate_input = st.number_input("$ Rate value", min_value=0.0, value=95.95, step=0.01, format="%.2f")

# mandatory = ["carat", "final", "vdb", "usa"]
# missing   = [next(l for kk, l in REQUIRED_COLS if kk == k)
#              for k in mandatory if not mapping.get(k)]
# if missing:
#     st.markdown(f'<div class="banner banner-warn">⚠️ Please map required columns: <b>{", ".join(missing)}</b></div>',
#                 unsafe_allow_html=True)
#     st.stop()

# # ─────────────────────────────────────────────────────────────────────────────
# #  STEP 3 – PREVIEW & DOWNLOAD
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="step-card">
#   <div class="step-header">
#     <span class="step-num">3</span>
#     <span class="step-title">Preview &amp; Download</span>
#   </div>
# </div>
# """, unsafe_allow_html=True)

# # ── Re-trim using the user's actual mapped columns (catches any edge cases) ──
# df = trim_data_rows(df, mapping.get("stock_no"), mapping.get("carat"))

# records = []
# for _, row in df.iterrows():
#     rec = {key: (row[col] if col else None) for key, col in mapping.items()}
#     records.append(rec)

# mapped_df = pd.DataFrame(records)
# for col in ["carat", "final", "vdb", "usa"]:
#     mapped_df[col] = pd.to_numeric(mapped_df[col], errors="coerce")

# mapped_df = mapped_df.dropna(subset=["carat", "final", "vdb", "usa"], how="all")
# n_stones  = len(mapped_df)

# if n_stones == 0:
#     st.markdown('<div class="banner banner-warn">⚠️ No valid rows found. Check your column mappings.</div>',
#                 unsafe_allow_html=True)
#     st.stop()

# # ── Arrange rows smallest → highest by column U (same order the U formula
# #    itself produces in Excel) ──────────────────────────────────────────────
# _sort_calc = add_calculated_fields(mapped_df, tariff_input, rate_input)
# mapped_df = (mapped_df.assign(_u=_sort_calc["U"].values)
#                        .sort_values("_u", kind="mergesort")
#                        .drop(columns="_u")
#                        .reset_index(drop=True))

# st.markdown(f'<div class="banner banner-success">✅ <b>{n_stones} stones</b> ready to export.</div>',
#             unsafe_allow_html=True)

# preview_rename = {
#     "stock_no":"Stock ID.", "certi_no":"Certi No.", "shape":"Shape",
#     "carat":"Carat", "lab":"Lab", "color":"Color", "clarity":"Clarity",
#     "final":"Final", "vdb":"VDB", "usa":"USA",
# }
# st.dataframe(mapped_df.rename(columns=preview_rename).head(10), use_container_width=True)

# with st.spinner("Building Excel with formulas and colours…"):
#     xlsx_bytes   = build_output_excel(mapped_df, rate_value=rate_input, tariff_value=tariff_input)
#     out_filename = safe_filename(party_file.name, n_stones)

# st.markdown('<div class="banner banner-success">💾 File ready — all formulas match the reference format exactly (cols P–Y yellow).</div>',
#             unsafe_allow_html=True)

# st.download_button(
#     label=f"⬇️  Download  {out_filename}",
#     data=xlsx_bytes,
#     file_name=out_filename,
#     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#     use_container_width=True,
# )

# with st.expander("📐 Formula & Colour Reference"):
#     st.markdown(f"""
# | Col | Header | Colour | Formula / Source |
# |---|---|---|---|
# | A | Stock ID | — | From party file |
# | B | Certi No | — | From party file |
# | C | Shape | — | From party file |
# | D | Carat | 🟦 Teal | From party file |
# | E | Lab | — | From party file |
# | F | Color | — | From party file |
# | G | Clarity | — | From party file |
# | H | Final | 🟨 Gold | From party file |
# | I | Amt | 🟨 Gold | `D × H` |
# | J | VDB | 🔵 Blue | From party file |
# | K | Amt | 🔵 Blue | `J × D` |
# | L | USA | 🟠 Orange | From party file |
# | M | Amt | 🟠 Orange | `L × D` |
# | N | With Tariff | 🟣 Lavender | `H × Tariff Multiplier` — Multiplier = **{tariff_input}** |
# | O | Amt | 🟣 Lavender | `N × D` |
# | P | — | 🟡 Yellow | `ROUND(((L-(M_avg−O_avg))/1.08),2)` |
# | Q | — | 🟡 Yellow | `ROUND(P×D, 2)` |
# | R | — | 🟡 Yellow | `ROUND(P×1.08, 2)` |
# | S | — | 🟡 Yellow | `ROUND(L−R, 2)` |
# | T | — | 🟡 Yellow | `ROUND((S×100)/P, 2)` |
# | U | — | 🟡 Yellow | `ROUND((L×ΣI) / SUMPRODUCT(ΣL,ΣD), 2)` |
# | V | — | 🟡 Yellow | `ROUND(U×D, 2)` |
# | W | — | 🟡 Yellow | `ROUND(L−U, 2)` |
# | X | — | 🟡 Yellow | `ROUND((W×100)/U, 2)` |
# | Y | — | 🟡 Yellow | `ROUND(U×$RATE, 2)` — RATE = **{rate_input}** |

# > 🟡 Both **TARIFF x** and **$ RATE** cells are highlighted yellow/red in the summary block — editable directly in Excel.
# > 🔽 Rows are sorted smallest → highest by column **U**, and an **AutoFilter (A:Z)** is enabled on the data table.
# """)

# import streamlit as st
# import pandas as pd
# import openpyxl
# from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
# from openpyxl.utils import get_column_letter
# import io
# import re
# import os
# from datetime import datetime

# # ─────────────────────────────────────────────────────────────────────────────
# #  PAGE CONFIG
# # ─────────────────────────────────────────────────────────────────────────────
# st.set_page_config(
#     page_title="💎 Diamond Formatter",
#     page_icon="💎",
#     layout="wide",
#     initial_sidebar_state="collapsed",
# )

# # ─────────────────────────────────────────────────────────────────────────────
# #  DARK THEME CSS
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <style>
# html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"] {
#     background-color: #0d0f14 !important; color: #e2e8f0 !important;
# }
# [data-testid="stHeader"] { background: #0d0f14 !important; }
# [data-testid="stSidebar"] { background: #111318 !important; }
# section[data-testid="stMain"] > div { background: #0d0f14 !important; }
# .block-container { padding-top: 1.5rem !important; }
# .hero {
#     position: relative;
#     background: linear-gradient(135deg, #0f1629 0%, #1a0a2e 50%, #0a1628 100%);
#     background-size: 200% 200%;
#     animation: heroGradientShift 12s ease infinite;
#     border: 1px solid #2a2d3e; border-radius: 20px;
#     padding: 2.4rem 2rem; margin-bottom: 2rem; overflow: hidden;
#     text-align: center; box-shadow: 0 10px 40px rgba(0,0,0,.35);
# }
# @keyframes heroGradientShift {
#     0%   { background-position: 0% 50%; }
#     50%  { background-position: 100% 50%; }
#     100% { background-position: 0% 50%; }
# }
# .hero::before, .hero::after {
#     content: ''; position: absolute; border-radius: 50%;
#     filter: blur(60px); opacity: .55;
#     animation: heroFloat 8s ease-in-out infinite; pointer-events: none;
# }
# .hero::before { width:240px;height:240px;background:rgba(99,102,241,.35);top:-70px;left:8%; }
# .hero::after  { width:280px;height:280px;background:rgba(168,85,247,.28);bottom:-90px;right:8%;animation-delay:-4s; }
# @keyframes heroFloat {
#     0%,100% { transform:translateY(0) translateX(0); }
#     50%     { transform:translateY(-22px) translateX(16px); }
# }
# .hero-content { position:relative;z-index:1;text-align:left; }
# .hero-company {
#     display:inline-block;font-size:.92rem;font-weight:700;letter-spacing:2.6px;
#     color:#c4b5fd;text-transform:uppercase;padding:7px 20px;
#     border:1px solid rgba(167,139,250,.45);border-radius:30px;
#     background:rgba(99,102,241,.12);margin:0 0 1.1rem;
#     animation:heroFadeInDown .8s ease both,heroPulseBorder 3s ease-in-out infinite 1s;
# }
# @keyframes heroFadeInDown { from{opacity:0;transform:translateY(-14px)} to{opacity:1;transform:translateY(0)} }
# @keyframes heroPulseBorder {
#     0%,100% { box-shadow:0 0 0 rgba(167,139,250,0); }
#     50%     { box-shadow:0 0 20px rgba(167,139,250,.4); }
# }
# .hero-title { display:flex;align-items:center;justify-content:flex-start;gap:12px;margin:0 0 .6rem;animation:heroFadeInUp .9s ease .2s both; }
# .hero-gem-wrap { position:relative;display:inline-flex;align-items:center;justify-content:center; }
# .hero-title .gem { display:inline-block;font-size:2rem;line-height:1;transform-style:preserve-3d;animation:heroGemSpin 4.5s ease-in-out infinite; }
# .hero-gem-wrap .spark { position:absolute;color:#7dd3fc;font-size:.6rem;opacity:0; }
# .hero-gem-wrap .spark-1 { top:-6px;right:-8px;animation:heroTwinkle 2.6s ease-in-out infinite; }
# .hero-gem-wrap .spark-2 { bottom:-4px;left:-10px;animation:heroTwinkle 2.6s ease-in-out infinite 1.1s; }
# @keyframes heroTwinkle { 0%,100%{opacity:0;transform:scale(.4) rotate(0deg)} 50%{opacity:1;transform:scale(1.1) rotate(25deg)} }
# @keyframes heroGemSpin {
#     0%   { transform:perspective(320px) rotateY(0deg) scale(1);filter:drop-shadow(0 0 6px rgba(167,139,250,.45)); }
#     45%  { transform:perspective(320px) rotateY(180deg) scale(1.1);filter:drop-shadow(0 0 16px rgba(56,189,248,.8)); }
#     55%  { transform:perspective(320px) rotateY(180deg) scale(1.1);filter:drop-shadow(0 0 16px rgba(56,189,248,.8)); }
#     100% { transform:perspective(320px) rotateY(360deg) scale(1);filter:drop-shadow(0 0 6px rgba(167,139,250,.45)); }
# }
# .hero-title .txt {
#     font-size:2.3rem;font-weight:800;letter-spacing:-.5px;line-height:1.15;
#     background:linear-gradient(90deg,#a78bfa,#6366f1,#38bdf8,#a78bfa);
#     background-size:300% auto;
#     -webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;
#     animation:heroShimmer 6s linear infinite;
# }
# @keyframes heroShimmer { to{background-position:300% center} }
# @keyframes heroFadeInUp { from{opacity:0;transform:translateY(18px)} to{opacity:1;transform:translateY(0)} }
# .hero-sub { color:#94a3b8;font-size:1rem;margin:0;animation:heroFadeInUp .9s ease .45s both; }
# .legend-wrap { display:flex;flex-wrap:wrap;gap:8px;margin:1rem 0 1.5rem; }
# .pill {
#     display:inline-flex;align-items:center;gap:7px;
#     background:#1a1d27;border:1px solid #2a2d3e;
#     border-radius:20px;padding:5px 12px;font-size:.78rem;color:#cbd5e1;
#     transition:transform .18s ease,border-color .18s ease;
# }
# .pill:hover { transform:translateY(-2px);border-color:#6366f1; }
# .dot { width:12px;height:12px;border-radius:3px;flex-shrink:0; }
# .step-card { background:#13161f;border:1px solid #1e2130;border-radius:12px;padding:1.5rem;margin-bottom:1rem; }
# .step-num {
#     display:inline-flex;align-items:center;justify-content:center;
#     width:32px;height:32px;border-radius:50%;
#     background:linear-gradient(135deg,#6366f1,#8b5cf6);
#     color:white;font-weight:700;font-size:.9rem;margin-right:10px;flex-shrink:0;
# }
# .step-title { font-size:1.1rem;font-weight:600;color:#f1f5f9; }
# .step-header { display:flex;align-items:center;margin-bottom:1rem; }
# [data-testid="stFileUploader"] { background:#13161f !important;border:2px dashed #2a2d3e !important;border-radius:10px !important; }
# [data-testid="stFileUploader"]:hover { border-color:#6366f1 !important; }
# .stSelectbox > div > div { background:#1a1d27 !important;border:1px solid #2a2d3e !important;color:#e2e8f0 !important;border-radius:8px !important; }
# .stSelectbox label { color:#cbd5e1 !important;font-size:.85rem !important; }
# .stDownloadButton button {
#     background:linear-gradient(135deg,#6366f1,#8b5cf6) !important;
#     color:white !important;border:none !important;border-radius:10px !important;
#     font-weight:600 !important;padding:.75rem 1.5rem !important;
#     font-size:1rem !important;transition:all .2s !important;
#     box-shadow:0 4px 20px rgba(99,102,241,.4) !important;
# }
# .stDownloadButton button:hover { transform:translateY(-2px) !important;box-shadow:0 8px 30px rgba(99,102,241,.6) !important; }
# .banner { border-radius:10px;padding:.85rem 1.1rem;font-size:.9rem;margin:.5rem 0;display:flex;align-items:center;gap:10px; }
# .banner-success { background:rgba(34,197,94,.1);border:1px solid rgba(34,197,94,.3);color:#86efac; }
# .banner-info    { background:rgba(99,102,241,.1);border:1px solid rgba(99,102,241,.3);color:#a5b4fc; }
# .banner-warn    { background:rgba(234,179,8,.1); border:1px solid rgba(234,179,8,.3); color:#fde68a; }
# .stats-row { display:flex;gap:12px;flex-wrap:wrap;margin:1rem 0; }
# .stat-box { background:#13161f;border:1px solid #1e2130;border-radius:10px;padding:.75rem 1.25rem;flex:1;min-width:130px; }
# .stat-val { font-size:1.6rem;font-weight:700;color:#a78bfa; }
# .stat-label { font-size:.75rem;color:#64748b;text-transform:uppercase;letter-spacing:.5px; }
# hr { border-color:#1e2130 !important;margin:1.5rem 0 !important; }
# [data-testid="stExpander"] { background:#13161f !important;border:1px solid #1e2130 !important;border-radius:10px !important; }
# summary { color:#94a3b8 !important; }
# </style>
# """, unsafe_allow_html=True)

# # ─────────────────────────────────────────────────────────────────────────────
# #  COLOR CONSTANTS
# #  Matches Image 2 exactly:
# #  - Cols A-O : existing group colors (navy header, teal, gold, blue, orange, lavender)
# #  - Cols P-Y : YELLOW (FFFF00) — boss requirement
# # ─────────────────────────────────────────────────────────────────────────────
# CLR_HEADER  = "1F4E79"   # deep navy header
# CLR_CARAT   = "98F6EB"   # teal   → col D (4)
# CLR_FINAL   = "FFCF37"   # gold   → H,I (8,9)
# CLR_VDB     = "5991D5"   # blue   → J,K (10,11)
# CLR_USA     = "F79B4F"   # orange → L,M (12,13)
# CLR_TARIFF  = "9E89B9"   # lavender → N,O (14,15)
# CLR_YELLOW  = "FFFF00"   # yellow → P-Y (16-25)

# # col index (1-based) → fill color for DATA cells
# COL_COLORS = {
#     4:  CLR_CARAT,
#     8:  CLR_FINAL,   9:  CLR_FINAL,
#     10: CLR_VDB,     11: CLR_VDB,
#     12: CLR_USA,     13: CLR_USA,
#     14: CLR_TARIFF,  15: CLR_TARIFF,
# }
# # cols 16-26 all yellow
# for _c in range(16, 27):
#     COL_COLORS[_c] = CLR_YELLOW

# # ─────────────────────────────────────────────────────────────────────────────
# #  OUTPUT HEADERS  (25 columns, matching Image 2)
# # ─────────────────────────────────────────────────────────────────────────────
# OUTPUT_HEADERS = [
#     "Stock ID.",        # A  1
#     "Certi No.",        # B  2
#     "Shape",            # C  3
#     "Carat",            # D  4
#     "Lab",              # E  5
#     "Color",            # F  6
#     "Clarity",          # G  7
#     "Final",            # H  8
#     "Amt",              # I  9
#     "VDB",              # J  10
#     "Amt",              # K  11
#     "USA",              # L  12
#     "Amt",              # M  13
#     "With Tariff",      # N  14
#     "Amt",              # O  15
#     "",                 # P  16  (yellow — unlabeled in source)
#     "",                 # Q  17
#     "",                 # R  18
#     "",                 # S  19
#     "",                 # T  20
#     "",                 # U  21
#     "",                 # V  22
#     "",                 # W  23
#     "",                 # X  24
#     "",                 # Y  25
#     "Total INR Amt",   # Z  26
# ]

# # NOTE: "with_tariff" is no longer taken from the party file — it is now
# # CALCULATED as Final × Tariff-Multiplier (user enters the multiplier, e.g.
# # 1.08 or 1.06), so it has been removed from the columns the user maps.
# REQUIRED_COLS = [
#     ("stock_no",    "Stock ID."),
#     ("certi_no",    "Certi No."),
#     ("shape",       "Shape"),
#     ("carat",       "Carat"),
#     ("lab",         "Lab"),
#     ("color",       "Color"),
#     ("clarity",     "Clarity"),
#     ("final",       "Final"),
#     ("vdb",         "VDB"),
#     ("usa",         "USA"),
# ]

# # ─────────────────────────────────────────────────────────────────────────────
# #  HELPERS
# # ─────────────────────────────────────────────────────────────────────────────
# def read_uploaded_file(f) -> pd.DataFrame | None:
#     name = f.name.lower()
#     try:
#         if name.endswith(".csv"):
#             return pd.read_csv(f)
#         elif name.endswith((".xlsx", ".xls", ".xlsm")):
#             return pd.read_excel(f, header=None)
#         elif name.endswith(".tsv"):
#             return pd.read_csv(f, sep="\t")
#         else:
#             st.error("Unsupported format. Please upload CSV, XLSX, XLS, XLSM, or TSV.")
#             return None
#     except Exception as e:
#         st.error(f"Could not read file: {e}")
#         return None


# def dedupe_columns(cols: list) -> list:
#     """
#     Some party files have duplicate header names (multiple 'AMT', 'Lab',
#     'Color', etc.). pandas/Arrow (used by st.dataframe) errors on duplicate
#     column names, so make every name unique: 'AMT', 'AMT_2', 'AMT_3', ...
#     """
#     seen: dict = {}
#     out = []
#     for c in cols:
#         name = str(c).strip() if c is not None and str(c).strip() else "Col"
#         if name not in seen:
#             seen[name] = 0
#             out.append(name)
#         else:
#             seen[name] += 1
#             out.append(f"{name}_{seen[name]}")
#     return out


# def detect_header_row(df: pd.DataFrame) -> int:
#     for i, row in df.iterrows():
#         vals = [str(v).strip() for v in row if pd.notna(v) and str(v).strip()]
#         if len(vals) >= 4 and not all(
#             re.sub(r'[.\-]', '', v).isdigit() for v in vals
#         ):
#             return i
#     return 0


# def trim_data_rows(df: pd.DataFrame, stock_col: str | None, carat_col: str | None) -> pd.DataFrame:
#     """
#     Drop any trailing non-data rows (SUM totals, labels, blanks) that appear
#     after the real diamond records.

#     A row is treated as a SUM / footer row — and all rows from it onward are
#     dropped — when ANY of these conditions holds:

#     1. Carat column is mapped AND the carat cell has a numeric value BUT the
#        stock-ID cell is empty or NaN.  This is the classic pattern:
#            Stock  | Carat
#            (blank)| 28.41   ← SUM row
#     2. Carat column is mapped AND the carat value is ≥ 3× the median carat of
#        the rows seen so far (catches a huge SUM value even if stock is filled).
#     3. Stock-ID column is mapped AND the stock-ID cell is a purely numeric string
#        or a known summary keyword (TOTAL, AVG, SUM, RATE, FINAL, etc.).
#     """
#     if df.empty:
#         return df

#     summary_keywords = {"total", "avg", "average", "sum", "rate", "final",
#                         "difference", "subtotal", "done", "fianl"}

#     # Compute median carat from first 5 valid rows for outlier detection
#     median_carat = None
#     if carat_col and carat_col in df.columns:
#         sample_carats = pd.to_numeric(df[carat_col].head(20), errors="coerce").dropna()
#         if not sample_carats.empty:
#             median_carat = sample_carats.median()

#     cut_at = len(df)  # default: keep everything

#     for i, (_, row) in enumerate(df.iterrows()):
#         # ── condition 1: empty stock + non-empty carat ──
#         if stock_col and stock_col in df.columns and carat_col and carat_col in df.columns:
#             stock_val = row.get(stock_col)
#             carat_val = row.get(carat_col)
#             stock_empty = pd.isna(stock_val) or str(stock_val).strip() == ""
#             carat_numeric = pd.notna(carat_val) and str(carat_val).strip() != ""
#             if stock_empty and carat_numeric:
#                 cut_at = i
#                 break

#         # ── condition 2: carat value is a huge outlier (SUM) ──
#         if carat_col and carat_col in df.columns and median_carat:
#             carat_val = pd.to_numeric(row.get(carat_col), errors="coerce")
#             if pd.notna(carat_val) and carat_val >= median_carat * 3:
#                 cut_at = i
#                 break

#         # ── condition 3: stock ID is a summary keyword ──
#         if stock_col and stock_col in df.columns:
#             stock_val = str(row.get(stock_col, "")).strip().lower()
#             if stock_val in summary_keywords:
#                 cut_at = i
#                 break

#     return df.iloc[:cut_at].reset_index(drop=True)


# def safe_filename(raw_name: str, n_stones: int) -> str:
#     """Party_Name_No_of_Stones_Current_Date.xlsx"""
#     stem = os.path.splitext(raw_name)[0]
#     stem = re.sub(r'[^\w\s\-]', '', stem).strip()
#     stem = re.sub(r'\s+', '_', stem)
#     date_str = datetime.now().strftime("%d-%m-%Y")
#     return f"{stem}_{n_stones}_Stones_{date_str}.xlsx"


# def smart_default(key: str, cols: list) -> int:
#     hints = {
#         "stock_no":    ["stock", "pexkt", "packet", "pkt", "id", "lot"],
#         "certi_no":    ["certi", "cert", "gia", "igi", "report"],
#         "shape":       ["shape"],
#         "carat":       ["carat", "ct", "weight"],
#         "lab":         ["lab", "igi", "gia"],
#         "color":       ["color", "colour"],
#         "clarity":     ["clarity", "clar"],
#         "final":       ["final", "rap%", "disc%", "back%"],
#         "vdb":         ["vdb"],
#         "usa":         ["usa"],
#     }
#     keywords = hints.get(key, [])
#     lower_cols = [c.lower() for c in cols[1:]]
#     for kw in keywords:
#         for i, lc in enumerate(lower_cols):
#             if kw in lc:
#                 return i + 1
#     return 0


# def px(hex_color: str) -> PatternFill:
#     return PatternFill("solid", fgColor=hex_color)


# def thin_border() -> Border:
#     s = Side(style="thin", color="D0D0D0")
#     return Border(left=s, right=s, top=s, bottom=s)


# def add_calculated_fields(df: pd.DataFrame, tariff_value: float, rate_value: float) -> pd.DataFrame:
#     """
#     Mirrors the exact Excel formulas (N, O, P..Z) in pandas so we can:
#       1) sort the data rows smallest → highest by column U before writing
#          them into the sheet (same order the U formula would produce), and
#       2) build the Summary sheet's insights/charts from the same numbers.
#     This does NOT change any Excel formula written by build_output_excel —
#     it's a read-only mirror used purely for ordering + reporting.
#     """
#     d = df.copy()
#     d["final_amt"]  = d["carat"] * d["final"]                 # I
#     d["vdb_amt"]    = d["vdb"] * d["carat"]                   # K
#     d["usa_amt"]    = d["usa"] * d["carat"]                   # M
#     d["tariff"]     = d["final"] * tariff_value                # N
#     d["tariff_amt"] = d["tariff"] * d["carat"]                 # O

#     sum_final_amt = d["final_amt"].sum()
#     sumproduct_ld = (d["usa"] * d["carat"]).sum()

#     d["U"] = (d["usa"] * sum_final_amt / sumproduct_ld) if sumproduct_ld else 0.0
#     d["V"] = d["U"] * d["carat"]
#     d["W"] = d["usa"] - d["U"]
#     d["X"] = (d["W"] * 100) / d["U"].replace(0, pd.NA)
#     d["Y"] = d["U"] * rate_value
#     d["total_inr"] = d["Y"] * d["carat"]                       # Z — Total INR Amt
#     return d


# # ─────────────────────────────────────────────────────────────────────────────
# #  CORE EXCEL BUILDER
# #
# #  Layout (matches Image 2 / source file exactly):
# #   Row 1        : Headers
# #   Rows 2..DS+n-1 : Data  (n = number of stones)
# #   SUM_ROW      : TOTAL
# #   AVG_ROW      : AVG / CARAT
# #   (blank rows)
# #   FINAL_ROW    : FINAL label + value
# #   TARIFF_ROW   : TARIFF x  label + value (yellow, red text) ← used by col-N formula
# #   RATE_ROW     : $ RATE label + value (yellow, red text) ← used by col-Y formula
# #   DIFF_ROW     : Difference
# #
# #  Column formulas (H onward) replicate source exactly, adjusted for output row refs:
# #   N  =ROUND(H*$N$TARIFF_ROW, 2)                 ← With Tariff = Final × Tariff multiplier
# #   O  =ROUND(N*D, 2)
# #   P  =ROUND(((L-(M_AVG - O_AVG))/1.08), 2)      ← where M_AVG=M$AVG_ROW, O_AVG=O$AVG_ROW
# #   Q  =ROUND(P*D, 2)
# #   R  =ROUND(P*1.08, 2)
# #   S  =ROUND(L-R, 2)
# #   T  =ROUND((S*100)/P, 2)
# #   U  =ROUND((L*$I$SUM_ROW)/SUMPRODUCT($L$DS:$L$DE,$D$DS:$D$DE), 2)
# #   V  =ROUND(U*D, 2)
# #   W  =ROUND(L-U, 2)
# #   X  =ROUND((W*100)/U, 2)
# #   Y  =ROUND(U*$N$RATE_ROW, 2)   ← references RATE cell instead of hardcoded 95.95
# #
# #  SUM row  : D, I, K, M, O, Q, V
# #  AVG row  : I, K, M, O, Q, V  (each SUM/D_SUM)
# # ─────────────────────────────────────────────────────────────────────────────
# def build_output_excel(mapped_df: pd.DataFrame, rate_value: float = 95.95,
#                         tariff_value: float = 1.08) -> bytes:
#     wb = openpyxl.Workbook()
#     ws = wb.active
#     ws.title = "Diamond Data"

#     n        = len(mapped_df)
#     HDR      = 1
#     DS       = 2            # data start row
#     DE       = DS + n - 1   # data end row
#     SUM_ROW    = DE + 1
#     AVG_ROW    = DE + 2
#     FINAL_ROW  = DE + 5
#     TARIFF_ROW = DE + 6
#     RATE_ROW   = DE + 7
#     DIFF_ROW   = DE + 8

#     # ── Header row ────────────────────────────────────────────────────────
#     for c, hdr in enumerate(OUTPUT_HEADERS, 1):
#         cell = ws.cell(row=HDR, column=c, value=hdr)
#         cell.fill      = px(CLR_HEADER)
#         cell.font      = Font(name="Arial", bold=True, color="FFFFFF", size=10)
#         cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
#         cell.border    = thin_border()
#     ws.row_dimensions[HDR].height = 36

#     # ── Data rows ─────────────────────────────────────────────────────────
#     for i, (_, row) in enumerate(mapped_df.iterrows()):
#         r = DS + i

#         # Raw input columns (A-M in output = cols 1-13, from party file)
#         raw = {
#             1:  row.get("stock_no"),
#             2:  row.get("certi_no"),
#             3:  row.get("shape"),
#             4:  row.get("carat"),
#             5:  row.get("lab"),
#             6:  row.get("color"),
#             7:  row.get("clarity"),
#             8:  row.get("final"),       # H  Final (%)
#             10: row.get("vdb"),         # J  VDB (%)
#             12: row.get("usa"),         # L  USA (%)
#         }

#         # Formula columns — col N (14) is now CALCULATED as Final × Tariff
#         # multiplier instead of being read from the party file.
#         formulas = {
#             9:  f"=ROUND(D{r}*H{r},2)",           # I  Final Amt
#             11: f"=ROUND(J{r}*D{r},2)",            # K  VDB Amt
#             13: f"=ROUND(L{r}*D{r},2)",            # M  USA Amt

#             # N  With Tariff = Final × Tariff multiplier (user-entered, cell ref)
#             14: f"=ROUND(H{r}*$N${TARIFF_ROW},2)",
#             # O  Tariff Amt = N*D
#             15: f"=ROUND(N{r}*D{r},2)",

#             # ── YELLOW columns (P-Y) — exact source formulas ──
#             # P: =ROUND(((L-(M_AVG-O_AVG))/1.08),2)
#             16: f"=ROUND(((L{r}-(M${AVG_ROW}-O${AVG_ROW}))/1.08),2)",
#             # Q: =ROUND(P*D,2)
#             17: f"=ROUND(P{r}*D{r},2)",
#             # R: =ROUND(P*1.08,2)
#             18: f"=ROUND(P{r}*1.08,2)",
#             # S: =ROUND(L-R,2)
#             19: f"=ROUND(L{r}-R{r},2)",
#             # T: =ROUND((S*100)/P,2)
#             20: f"=ROUND((S{r}*100)/P{r},2)",
#             # U: =ROUND((L*$I$SUM)/SUMPRODUCT($L$DS:$L$DE,$D$DS:$D$DE),2)
#             21: f"=ROUND((L{r}*$I${SUM_ROW})/SUMPRODUCT($L${DS}:$L${DE},$D${DS}:$D${DE}),2)",
#             # V: =ROUND(U*D,2)
#             22: f"=ROUND(U{r}*D{r},2)",
#             # W: =ROUND(L-U,2)
#             23: f"=ROUND(L{r}-U{r},2)",
#             # X: =ROUND((W*100)/U,2)
#             24: f"=ROUND((W{r}*100)/U{r},2)",
#             # Y: =ROUND(U*RATE,2)  — references the RATE cell
#             25: f"=ROUND(U{r}*$N${RATE_ROW},2)",
#             # Z: =ROUND(Y*D,2)  — Total INR Amt
#             26: f"=ROUND(Y{r}*D{r},2)",
#         }

#         for c in range(1, 27):
#             val = raw.get(c, formulas.get(c))
#             cell = ws.cell(row=r, column=c, value=val)
#             cell.border    = thin_border()
#             cell.alignment = Alignment(horizontal="center", vertical="center")
#             cell.font      = Font(name="Arial", size=10)
#             if c in COL_COLORS:
#                 cell.fill = px(COL_COLORS[c])
#             if c in (4, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26):
#                 cell.number_format = "0.00"

#         ws.row_dimensions[r].height = 18

#     # ── TOTAL row ─────────────────────────────────────────────────────────
#     total_fill = px("F2F2F2")
#     lbl = ws.cell(row=SUM_ROW, column=1, value="TOTAL")
#     lbl.font = Font(name="Arial", bold=True, size=10)
#     lbl.fill = total_fill; lbl.border = thin_border()
#     lbl.alignment = Alignment(horizontal="center")

#     # SUM cols: D(4), I(9), K(11), M(13), O(15), Q(17), V(22), Z(26)
#     sum_cols = {4:"D", 9:"I", 11:"K", 13:"M", 15:"O", 17:"Q", 22:"V", 26:"Z"}
#     for col, letter in sum_cols.items():
#         cell = ws.cell(row=SUM_ROW, column=col,
#                        value=f"=SUM({letter}{DS}:{letter}{DE})")
#         cell.font = Font(name="Arial", bold=True, size=10)
#         cell.fill = total_fill; cell.border = thin_border()
#         cell.alignment = Alignment(horizontal="center")
#         cell.number_format = "0.00"

#     # ── AVG / CARAT row ───────────────────────────────────────────────────
#     avg_fill = px("E2EFDA")
#     lbl2 = ws.cell(row=AVG_ROW, column=1, value="AVG / CARAT")
#     lbl2.font = Font(name="Arial", bold=True, size=10)
#     lbl2.fill = avg_fill; lbl2.border = thin_border()
#     lbl2.alignment = Alignment(horizontal="center")

#     # AVG cols: I(9), K(11), M(13), O(15), Q(17), V(22)
#     avg_cols = {9:"I", 11:"K", 13:"M", 15:"O", 17:"Q", 22:"V"}
#     for col, letter in avg_cols.items():
#         cell = ws.cell(row=AVG_ROW, column=col,
#                        value=f"=ROUND({letter}{SUM_ROW}/D{SUM_ROW},2)")
#         cell.font = Font(name="Arial", bold=True, size=10)
#         cell.fill = avg_fill; cell.border = thin_border()
#         cell.alignment = Alignment(horizontal="center")
#         cell.number_format = "0.00"

#     # ── FINAL / TARIFF / RATE / Difference label block ───────────────────
#     # FINAL row
#     for c, val in ((13, "FINAL"), (14, f"=ROUND(I{AVG_ROW},2)")):
#         cell = ws.cell(row=FINAL_ROW, column=c, value=val)
#         cell.font = Font(name="Arial", bold=True, size=10)
#         cell.border = thin_border()
#         cell.alignment = Alignment(horizontal="center")
#         cell.number_format = "0.00"

#     # TARIFF row — label M, value N (yellow fill, red bold text for value)
#     # This is the multiplier used to CALCULATE "With Tariff" (col N) as
#     # Final × Tariff, e.g. 1.08 or 1.06 — editable directly in Excel.
#     lbl_tariff = ws.cell(row=TARIFF_ROW, column=13, value="TARIFF x")
#     lbl_tariff.font = Font(name="Arial", bold=True, size=10)
#     lbl_tariff.border = thin_border()
#     lbl_tariff.alignment = Alignment(horizontal="center")

#     tariff_cell = ws.cell(row=TARIFF_ROW, column=14, value=tariff_value)
#     tariff_cell.fill   = px("FFFF00")
#     tariff_cell.font   = Font(name="Arial", bold=True, color="FF0000", size=10)
#     tariff_cell.border = thin_border()
#     tariff_cell.alignment = Alignment(horizontal="center")
#     tariff_cell.number_format = "0.00"

#     # RATE row — label M, value N (yellow fill, red bold text for value)
#     lbl_rate = ws.cell(row=RATE_ROW, column=13, value="$ RATE")
#     lbl_rate.font = Font(name="Arial", bold=True, size=10)
#     lbl_rate.border = thin_border()
#     lbl_rate.alignment = Alignment(horizontal="center")

#     rate_cell = ws.cell(row=RATE_ROW, column=14, value=rate_value)
#     rate_cell.fill   = px("FFFF00")
#     rate_cell.font   = Font(name="Arial", bold=True, color="FF0000", size=10)
#     rate_cell.border = thin_border()
#     rate_cell.alignment = Alignment(horizontal="center")
#     rate_cell.number_format = "0.00"

#     # Difference row
#     for c, val in ((13, "Difference"), (14, f"=M{AVG_ROW}-O{AVG_ROW}")):
#         cell = ws.cell(row=DIFF_ROW, column=c, value=val)
#         cell.font = Font(name="Arial", bold=True, size=10)
#         cell.border = thin_border()
#         cell.alignment = Alignment(horizontal="center")
#         cell.number_format = "0.00"

#     # ── Column widths ──────────────────────────────────────────────────────
#     widths = [14, 14, 13, 7, 6, 7, 9,  8, 11,  8, 11,  8, 11, 12, 11,
#               11, 11, 11,  8,  9, 11, 11, 11,  9, 13, 15]
#     for i, w in enumerate(widths, 1):
#         ws.column_dimensions[get_column_letter(i)].width = w

#     ws.freeze_panes = "A2"

#     # ── AutoFilter across the header + data table (A:Z) ─────────────────────
#     ws.auto_filter.ref = f"A{HDR}:Z{DE}"

#     buf = io.BytesIO()
#     wb.save(buf)
#     buf.seek(0)
#     return buf.getvalue()


# # ─────────────────────────────────────────────────────────────────────────────
# #  HERO BANNER
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="hero">
#   <div class="hero-content">
#     <p class="hero-company">✦ BRIGHT TRADING INDIA LLP</p>
#     <p class="hero-title">
#       <span class="hero-gem-wrap">
#         <span class="gem">💎</span>
#         <span class="spark spark-1">✦</span>
#         <span class="spark spark-2">✦</span>
#       </span>
#       <span class="txt">Diamond Cost Optimizer</span>
#     </p>
#     <p class="hero-sub">Smart Cost Adjustment &amp; Equal Profit Calculator</p>
#   </div>
# </div>
# """, unsafe_allow_html=True)

# st.markdown("""
# <div class="legend-wrap">
#   <div class="pill"><div class="dot" style="background:#98F6EB"></div>Carat</div>
#   <div class="pill"><div class="dot" style="background:#FFCF37"></div>Final / Amt</div>
#   <div class="pill"><div class="dot" style="background:#5991D5"></div>VDB / Amt</div>
#   <div class="pill"><div class="dot" style="background:#F79B4F"></div>USA / Amt</div>
#   <div class="pill"><div class="dot" style="background:#9E89B9"></div>With Tariff / Amt</div>
#   <div class="pill"><div class="dot" style="background:#FFFF00"></div>Calculated Columns (P–Y)</div>
# </div>
# """, unsafe_allow_html=True)

# # ─────────────────────────────────────────────────────────────────────────────
# #  STEP 1 – UPLOAD
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="step-card">
#   <div class="step-header">
#     <span class="step-num">1</span>
#     <span class="step-title">Upload Party File</span>
#   </div>
#   <p style="color:#64748b;font-size:.87rem;margin:0 0 .9rem 42px;">
#     Accepts <b style="color:#a78bfa">CSV, XLSX, XLS, XLSM, TSV</b> — any column order, any layout.
#   </p>
# </div>
# """, unsafe_allow_html=True)

# party_file = st.file_uploader(
#     "Drop your file here or click Upload",
#     type=["csv", "xlsx", "xls", "xlsm", "tsv"],
#     key="party_file",
#     label_visibility="collapsed",
# )

# if party_file is None:
#     st.markdown('<div class="banner banner-info">👆 Upload the party file above to begin.</div>',
#                 unsafe_allow_html=True)
#     st.stop()

# raw_df = read_uploaded_file(party_file)
# if raw_df is None:
#     st.stop()

# h_row = detect_header_row(raw_df)
# df = raw_df.copy()
# df.columns = raw_df.iloc[h_row]
# df = df.iloc[h_row + 1:].reset_index(drop=True)
# df.columns = dedupe_columns([str(c).strip() if pd.notna(c) else f"Col_{i}"
#                               for i, c in enumerate(df.columns)])
# df = df.dropna(how="all")

# # ── Auto-detect and strip SUM/footer rows at the bottom ──────────────────────
# # We do a quick keyword-based column sniff so trim works before Step-2 mapping
# def _sniff_col(df, keywords):
#     for col in df.columns:
#         if any(kw in str(col).lower() for kw in keywords):
#             return col
#     return None

# _stock_col = _sniff_col(df, ["stock", "packet", "pkt", "lot", "id"])
# _carat_col = _sniff_col(df, ["carat", "ct", "weight"])
# df = trim_data_rows(df, _stock_col, _carat_col)

# st.markdown(f"""
# <div class="stats-row">
#   <div class="stat-box"><div class="stat-val">{len(df)}</div><div class="stat-label">Rows Detected</div></div>
#   <div class="stat-box"><div class="stat-val">{len(df.columns)}</div><div class="stat-label">Columns Found</div></div>
#   <div class="stat-box"><div class="stat-val">{party_file.name.rsplit(".",1)[-1].upper()}</div><div class="stat-label">File Format</div></div>
# </div>
# """, unsafe_allow_html=True)

# with st.expander("👀 Preview raw data (first 5 rows)"):
#     st.dataframe(df.head(5), use_container_width=True)

# # ─────────────────────────────────────────────────────────────────────────────
# #  STEP 2 – COLUMN MAPPING
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="step-card">
#   <div class="step-header">
#     <span class="step-num">2</span>
#     <span class="step-title">Map Columns</span>
#   </div>
# </div>
# """, unsafe_allow_html=True)
# st.caption("Select which column in the party file corresponds to each required field.")

# available = ["— skip / not available —"] + list(df.columns)
# mapping   = {}

# cols_left, cols_right = st.columns(2)
# for idx, (key, label) in enumerate(REQUIRED_COLS):
#     default = smart_default(key, available)
#     container = cols_left if idx % 2 == 0 else cols_right
#     with container:
#         choice = st.selectbox(f"**{label}**", options=available,
#                               index=default, key=f"map_{key}")
#         mapping[key] = None if choice.startswith("—") else choice

# # ── Tariff multiplier + $ RATE inputs ─────────────────────────────────────────
# st.markdown("---")
# # st.markdown("**Tariff Multiplier** — used to calculate the `With Tariff` column (`Final × Multiplier`), e.g. 1.08 or 1.06")
# tariff_input = st.number_input("Tariff multiplier", min_value=1.0, value=1.08, step=0.01, format="%.2f")

# # st.markdown("**$ Rate** — used in the final Y-column formula (`U × $ Rate`)")
# rate_input = st.number_input("$ Rate value", min_value=0.0, value=95.95, step=0.01, format="%.2f")

# mandatory = ["carat", "final", "vdb", "usa"]
# missing   = [next(l for kk, l in REQUIRED_COLS if kk == k)
#              for k in mandatory if not mapping.get(k)]
# if missing:
#     st.markdown(f'<div class="banner banner-warn">⚠️ Please map required columns: <b>{", ".join(missing)}</b></div>',
#                 unsafe_allow_html=True)
#     st.stop()

# # ─────────────────────────────────────────────────────────────────────────────
# #  STEP 3 – PREVIEW & DOWNLOAD
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="step-card">
#   <div class="step-header">
#     <span class="step-num">3</span>
#     <span class="step-title">Preview &amp; Download</span>
#   </div>
# </div>
# """, unsafe_allow_html=True)

# # ── Re-trim using the user's actual mapped columns (catches any edge cases) ──
# df = trim_data_rows(df, mapping.get("stock_no"), mapping.get("carat"))

# records = []
# for _, row in df.iterrows():
#     rec = {key: (row[col] if col else None) for key, col in mapping.items()}
#     records.append(rec)

# mapped_df = pd.DataFrame(records)
# for col in ["carat", "final", "vdb", "usa"]:
#     mapped_df[col] = pd.to_numeric(mapped_df[col], errors="coerce")

# mapped_df = mapped_df.dropna(subset=["carat", "final", "vdb", "usa"], how="all")
# n_stones  = len(mapped_df)

# if n_stones == 0:
#     st.markdown('<div class="banner banner-warn">⚠️ No valid rows found. Check your column mappings.</div>',
#                 unsafe_allow_html=True)
#     st.stop()

# # ── Arrange rows smallest → highest by column U (same order the U formula
# #    itself produces in Excel) ──────────────────────────────────────────────
# _sort_calc = add_calculated_fields(mapped_df, tariff_input, rate_input)
# mapped_df = (mapped_df.assign(_u=_sort_calc["U"].values)
#                        .sort_values("_u", kind="mergesort")
#                        .drop(columns="_u")
#                        .reset_index(drop=True))

# st.markdown(f'<div class="banner banner-success">✅ <b>{n_stones} stones</b> ready to export.</div>',
#             unsafe_allow_html=True)

# preview_rename = {
#     "stock_no":"Stock ID.", "certi_no":"Certi No.", "shape":"Shape",
#     "carat":"Carat", "lab":"Lab", "color":"Color", "clarity":"Clarity",
#     "final":"Final", "vdb":"VDB", "usa":"USA",
# }
# st.dataframe(mapped_df.rename(columns=preview_rename).head(10), use_container_width=True)

# with st.spinner("Building Excel with formulas and colours…"):
#     xlsx_bytes   = build_output_excel(mapped_df, rate_value=rate_input, tariff_value=tariff_input)
#     out_filename = safe_filename(party_file.name, n_stones)

# st.markdown('<div class="banner banner-success">💾 File ready — all formulas match the reference format exactly (cols P–Y yellow).</div>',
#             unsafe_allow_html=True)

# st.download_button(
#     label=f"⬇️  Download  {out_filename}",
#     data=xlsx_bytes,
#     file_name=out_filename,
#     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#     use_container_width=True,
# )

# with st.expander("📐 Formula & Colour Reference"):
#     st.markdown(f"""
# | Col | Header | Colour | Formula / Source |
# |---|---|---|---|
# | A | Stock ID | — | From party file |
# | B | Certi No | — | From party file |
# | C | Shape | — | From party file |
# | D | Carat | 🟦 Teal | From party file |
# | E | Lab | — | From party file |
# | F | Color | — | From party file |
# | G | Clarity | — | From party file |
# | H | Final | 🟨 Gold | From party file |
# | I | Amt | 🟨 Gold | `D × H` |
# | J | VDB | 🔵 Blue | From party file |
# | K | Amt | 🔵 Blue | `J × D` |
# | L | USA | 🟠 Orange | From party file |
# | M | Amt | 🟠 Orange | `L × D` |
# | N | With Tariff | 🟣 Lavender | `H × Tariff Multiplier` — Multiplier = **{tariff_input}** |
# | O | Amt | 🟣 Lavender | `N × D` |
# | P | — | 🟡 Yellow | `ROUND(((L-(M_avg−O_avg))/1.08),2)` |
# | Q | — | 🟡 Yellow | `ROUND(P×D, 2)` |
# | R | — | 🟡 Yellow | `ROUND(P×1.08, 2)` |
# | S | — | 🟡 Yellow | `ROUND(L−R, 2)` |
# | T | — | 🟡 Yellow | `ROUND((S×100)/P, 2)` |
# | U | — | 🟡 Yellow | `ROUND((L×ΣI) / SUMPRODUCT(ΣL,ΣD), 2)` |
# | V | — | 🟡 Yellow | `ROUND(U×D, 2)` |
# | W | — | 🟡 Yellow | `ROUND(L−U, 2)` |
# | X | — | 🟡 Yellow | `ROUND((W×100)/U, 2)` |
# | Y | — | 🟡 Yellow | `ROUND(U×$RATE, 2)` — RATE = **{rate_input}** |

# > 🟡 Both **TARIFF x** and **$ RATE** cells are highlighted yellow/red in the summary block — editable directly in Excel.
# > 🔽 Rows are sorted smallest → highest by column **U**, and an **AutoFilter (A:Z)** is enabled on the data table.
# """)











# ALL COMBINE










# import streamlit as st
# import pandas as pd
# import openpyxl
# from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
# from openpyxl.utils import get_column_letter
# import io
# import re
# import os
# from datetime import datetime

# # ─────────────────────────────────────────────────────────────────────────────
# #  PAGE CONFIG
# # ─────────────────────────────────────────────────────────────────────────────
# st.set_page_config(
#     page_title="💎 Diamond Formatter",
#     page_icon="💎",
#     layout="wide",
#     initial_sidebar_state="collapsed",
# )

# # ─────────────────────────────────────────────────────────────────────────────
# #  DARK THEME CSS
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <style>
# html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"] {
#     background-color: #0d0f14 !important; color: #e2e8f0 !important;
# }
# [data-testid="stHeader"] { background: #0d0f14 !important; }
# [data-testid="stSidebar"] { background: #111318 !important; }
# section[data-testid="stMain"] > div { background: #0d0f14 !important; }
# .block-container { padding-top: 1.5rem !important; }
# .hero {
#     position: relative;
#     background: linear-gradient(135deg, #0f1629 0%, #1a0a2e 50%, #0a1628 100%);
#     background-size: 200% 200%;
#     animation: heroGradientShift 12s ease infinite;
#     border: 1px solid #2a2d3e; border-radius: 20px;
#     padding: 2.4rem 2rem; margin-bottom: 2rem; overflow: hidden;
#     text-align: center; box-shadow: 0 10px 40px rgba(0,0,0,.35);
# }
# @keyframes heroGradientShift {
#     0%   { background-position: 0% 50%; }
#     50%  { background-position: 100% 50%; }
#     100% { background-position: 0% 50%; }
# }
# .hero::before, .hero::after {
#     content: ''; position: absolute; border-radius: 50%;
#     filter: blur(60px); opacity: .55;
#     animation: heroFloat 8s ease-in-out infinite; pointer-events: none;
# }
# .hero::before { width:240px;height:240px;background:rgba(99,102,241,.35);top:-70px;left:8%; }
# .hero::after  { width:280px;height:280px;background:rgba(168,85,247,.28);bottom:-90px;right:8%;animation-delay:-4s; }
# @keyframes heroFloat {
#     0%,100% { transform:translateY(0) translateX(0); }
#     50%     { transform:translateY(-22px) translateX(16px); }
# }
# .hero-content { position:relative;z-index:1;text-align:left; }
# .hero-company {
#     display:inline-block;font-size:.92rem;font-weight:700;letter-spacing:2.6px;
#     color:#c4b5fd;text-transform:uppercase;padding:7px 20px;
#     border:1px solid rgba(167,139,250,.45);border-radius:30px;
#     background:rgba(99,102,241,.12);margin:0 0 1.1rem;
#     animation:heroFadeInDown .8s ease both,heroPulseBorder 3s ease-in-out infinite 1s;
# }
# @keyframes heroFadeInDown { from{opacity:0;transform:translateY(-14px)} to{opacity:1;transform:translateY(0)} }
# @keyframes heroPulseBorder {
#     0%,100% { box-shadow:0 0 0 rgba(167,139,250,0); }
#     50%     { box-shadow:0 0 20px rgba(167,139,250,.4); }
# }
# .hero-title { display:flex;align-items:center;justify-content:flex-start;gap:12px;margin:0 0 .6rem;animation:heroFadeInUp .9s ease .2s both; }
# .hero-gem-wrap { position:relative;display:inline-flex;align-items:center;justify-content:center; }
# .hero-title .gem { display:inline-block;font-size:2rem;line-height:1;transform-style:preserve-3d;animation:heroGemSpin 4.5s ease-in-out infinite; }
# .hero-gem-wrap .spark { position:absolute;color:#7dd3fc;font-size:.6rem;opacity:0; }
# .hero-gem-wrap .spark-1 { top:-6px;right:-8px;animation:heroTwinkle 2.6s ease-in-out infinite; }
# .hero-gem-wrap .spark-2 { bottom:-4px;left:-10px;animation:heroTwinkle 2.6s ease-in-out infinite 1.1s; }
# @keyframes heroTwinkle { 0%,100%{opacity:0;transform:scale(.4) rotate(0deg)} 50%{opacity:1;transform:scale(1.1) rotate(25deg)} }
# @keyframes heroGemSpin {
#     0%   { transform:perspective(320px) rotateY(0deg) scale(1);filter:drop-shadow(0 0 6px rgba(167,139,250,.45)); }
#     45%  { transform:perspective(320px) rotateY(180deg) scale(1.1);filter:drop-shadow(0 0 16px rgba(56,189,248,.8)); }
#     55%  { transform:perspective(320px) rotateY(180deg) scale(1.1);filter:drop-shadow(0 0 16px rgba(56,189,248,.8)); }
#     100% { transform:perspective(320px) rotateY(360deg) scale(1);filter:drop-shadow(0 0 6px rgba(167,139,250,.45)); }
# }
# .hero-title .txt {
#     font-size:2.3rem;font-weight:800;letter-spacing:-.5px;line-height:1.15;
#     background:linear-gradient(90deg,#a78bfa,#6366f1,#38bdf8,#a78bfa);
#     background-size:300% auto;
#     -webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;
#     animation:heroShimmer 6s linear infinite;
# }
# @keyframes heroShimmer { to{background-position:300% center} }
# @keyframes heroFadeInUp { from{opacity:0;transform:translateY(18px)} to{opacity:1;transform:translateY(0)} }
# .hero-sub { color:#94a3b8;font-size:1rem;margin:0;animation:heroFadeInUp .9s ease .45s both; }
# .legend-wrap { display:flex;flex-wrap:wrap;gap:8px;margin:1rem 0 1.5rem; }
# .pill {
#     display:inline-flex;align-items:center;gap:7px;
#     background:#1a1d27;border:1px solid #2a2d3e;
#     border-radius:20px;padding:5px 12px;font-size:.78rem;color:#cbd5e1;
#     transition:transform .18s ease,border-color .18s ease;
# }
# .pill:hover { transform:translateY(-2px);border-color:#6366f1; }
# .dot { width:12px;height:12px;border-radius:3px;flex-shrink:0; }
# .step-card { background:#13161f;border:1px solid #1e2130;border-radius:12px;padding:1.5rem;margin-bottom:1rem; }
# .step-num {
#     display:inline-flex;align-items:center;justify-content:center;
#     width:32px;height:32px;border-radius:50%;
#     background:linear-gradient(135deg,#6366f1,#8b5cf6);
#     color:white;font-weight:700;font-size:.9rem;margin-right:10px;flex-shrink:0;
# }
# .step-title { font-size:1.1rem;font-weight:600;color:#f1f5f9; }
# .step-header { display:flex;align-items:center;margin-bottom:1rem; }
# [data-testid="stFileUploader"] { background:#13161f !important;border:2px dashed #2a2d3e !important;border-radius:10px !important; }
# [data-testid="stFileUploader"]:hover { border-color:#6366f1 !important; }
# .stSelectbox > div > div { background:#1a1d27 !important;border:1px solid #2a2d3e !important;color:#e2e8f0 !important;border-radius:8px !important; }
# .stSelectbox label { color:#cbd5e1 !important;font-size:.85rem !important; }
# .stDownloadButton button {
#     background:linear-gradient(135deg,#6366f1,#8b5cf6) !important;
#     color:white !important;border:none !important;border-radius:10px !important;
#     font-weight:600 !important;padding:.75rem 1.5rem !important;
#     font-size:1rem !important;transition:all .2s !important;
#     box-shadow:0 4px 20px rgba(99,102,241,.4) !important;
# }
# .stDownloadButton button:hover { transform:translateY(-2px) !important;box-shadow:0 8px 30px rgba(99,102,241,.6) !important; }
# .banner { border-radius:10px;padding:.85rem 1.1rem;font-size:.9rem;margin:.5rem 0;display:flex;align-items:center;gap:10px; }
# .banner-success { background:rgba(34,197,94,.1);border:1px solid rgba(34,197,94,.3);color:#86efac; }
# .banner-info    { background:rgba(99,102,241,.1);border:1px solid rgba(99,102,241,.3);color:#a5b4fc; }
# .banner-warn    { background:rgba(234,179,8,.1); border:1px solid rgba(234,179,8,.3); color:#fde68a; }
# .stats-row { display:flex;gap:12px;flex-wrap:wrap;margin:1rem 0; }
# .stat-box { background:#13161f;border:1px solid #1e2130;border-radius:10px;padding:.75rem 1.25rem;flex:1;min-width:130px; }
# .stat-val { font-size:1.6rem;font-weight:700;color:#a78bfa; }
# .stat-label { font-size:.75rem;color:#64748b;text-transform:uppercase;letter-spacing:.5px; }
# hr { border-color:#1e2130 !important;margin:1.5rem 0 !important; }
# [data-testid="stExpander"] { background:#13161f !important;border:1px solid #1e2130 !important;border-radius:10px !important; }
# summary { color:#94a3b8 !important; }
# </style>
# """, unsafe_allow_html=True)

# # ─────────────────────────────────────────────────────────────────────────────
# #  COLOR CONSTANTS
# #  Matches Image 2 exactly:
# #  - Cols A-O : existing group colors (navy header, teal, gold, blue, orange, lavender)
# #  - Cols P-Y : YELLOW (FFFF00) — boss requirement
# # ─────────────────────────────────────────────────────────────────────────────
# CLR_HEADER  = "1F4E79"   # deep navy header
# CLR_CARAT   = "98F6EB"   # teal   → col D (4)
# CLR_FINAL   = "FFCF37"   # gold   → H,I (8,9)
# CLR_VDB     = "5991D5"   # blue   → J,K (10,11)
# CLR_USA     = "F79B4F"   # orange → L,M (12,13)
# CLR_TARIFF  = "9E89B9"   # lavender → N,O (14,15)
# CLR_YELLOW  = "FFFF00"   # yellow → P-Y (16-25)

# # col index (1-based) → fill color for DATA cells
# COL_COLORS = {
#     4:  CLR_CARAT,
#     8:  CLR_FINAL,   9:  CLR_FINAL,
#     10: CLR_VDB,     11: CLR_VDB,
#     12: CLR_USA,     13: CLR_USA,
#     14: CLR_TARIFF,  15: CLR_TARIFF,
# }
# # cols 16-26 all yellow
# for _c in range(16, 27):
#     COL_COLORS[_c] = CLR_YELLOW

# # ─────────────────────────────────────────────────────────────────────────────
# #  OUTPUT HEADERS  (25 columns, matching Image 2)
# # ─────────────────────────────────────────────────────────────────────────────
# OUTPUT_HEADERS = [
#     "Stock ID.",        # A  1
#     "Certi No.",        # B  2
#     "Shape",            # C  3
#     "Carat",            # D  4
#     "Lab",              # E  5
#     "Color",            # F  6
#     "Clarity",          # G  7
#     "Final",            # H  8
#     "Amt",              # I  9
#     "VDB",              # J  10
#     "Amt",              # K  11
#     "USA",              # L  12
#     "Amt",              # M  13
#     "With Tariff",      # N  14
#     "Amt",              # O  15
#     "",                 # P  16  (yellow — unlabeled in source)
#     "",                 # Q  17
#     "",                 # R  18
#     "",                 # S  19
#     "",                 # T  20
#     "",                 # U  21
#     "",                 # V  22
#     "",                 # W  23
#     "",                 # X  24
#     "",                 # Y  25
#     "Total INR Amt",   # Z  26
# ]

# # NOTE: "with_tariff" is no longer taken from the party file — it is now
# # CALCULATED as Final × Tariff-Multiplier (user enters the multiplier, e.g.
# # 1.08 or 1.06), so it has been removed from the columns the user maps.
# REQUIRED_COLS = [
#     ("stock_no",    "Stock ID."),
#     ("certi_no",    "Certi No."),
#     ("shape",       "Shape"),
#     ("carat",       "Carat"),
#     ("lab",         "Lab"),
#     ("color",       "Color"),
#     ("clarity",     "Clarity"),
#     ("final",       "Final"),
#     ("vdb",         "VDB"),
#     ("usa",         "USA"),
# ]

# # ─────────────────────────────────────────────────────────────────────────────
# #  HELPERS
# # ─────────────────────────────────────────────────────────────────────────────
# def read_uploaded_file(f) -> pd.DataFrame | None:
#     name = f.name.lower()
#     try:
#         if name.endswith(".csv"):
#             return pd.read_csv(f)
#         elif name.endswith((".xlsx", ".xls", ".xlsm")):
#             return pd.read_excel(f, header=None)
#         elif name.endswith(".tsv"):
#             return pd.read_csv(f, sep="\t")
#         else:
#             st.error("Unsupported format. Please upload CSV, XLSX, XLS, XLSM, or TSV.")
#             return None
#     except Exception as e:
#         st.error(f"Could not read file: {e}")
#         return None


# def dedupe_columns(cols: list) -> list:
#     """
#     Some party files have duplicate header names (multiple 'AMT', 'Lab',
#     'Color', etc.). pandas/Arrow (used by st.dataframe) errors on duplicate
#     column names, so make every name unique: 'AMT', 'AMT_2', 'AMT_3', ...
#     """
#     seen: dict = {}
#     out = []
#     for c in cols:
#         name = str(c).strip() if c is not None and str(c).strip() else "Col"
#         if name not in seen:
#             seen[name] = 0
#             out.append(name)
#         else:
#             seen[name] += 1
#             out.append(f"{name}_{seen[name]}")
#     return out


# def detect_header_row(df: pd.DataFrame) -> int:
#     for i, row in df.iterrows():
#         vals = [str(v).strip() for v in row if pd.notna(v) and str(v).strip()]
#         if len(vals) >= 4 and not all(
#             re.sub(r'[.\-]', '', v).isdigit() for v in vals
#         ):
#             return i
#     return 0


# def trim_data_rows(df: pd.DataFrame, stock_col: str | None, carat_col: str | None) -> pd.DataFrame:
#     """
#     Remove SUM / TOTAL / AVG / footer rows from anywhere in the sheet — not
#     just a trailing block. Some party files have a summary block in the
#     MIDDLE of the sheet (e.g. 23 stones → TOTAL/AVG rows → 4 more stones
#     below that). The old version cut everything after the first junk row it
#     saw, which silently dropped those extra genuine stones. Now each row is
#     judged independently, so real stone rows below a summary block are kept.

#     A row is treated as junk (dropped) when ANY of these hold:

#     1. Carat column is mapped AND the carat cell has a numeric value BUT the
#        stock-ID cell is empty or NaN. This is the classic pattern:
#            Stock  | Carat
#            (blank)| 28.41   ← SUM row
#     2. Carat column is mapped AND the carat value is ≥ 3× the median carat of
#        genuine stone rows (rows with a non-empty stock ID) — catches a huge
#        SUM value even if stock happens to be filled.
#     3. Stock-ID column is mapped AND the stock-ID cell is a known summary
#        keyword (TOTAL, AVG, SUM, RATE, FINAL, etc.).
#     """
#     if df.empty:
#         return df

#     summary_keywords = {"total", "avg", "average", "sum", "rate", "final",
#                         "difference", "subtotal", "done", "fianl"}

#     # Estimate median carat from rows that look like real stones (non-empty
#     # stock ID) across the WHOLE sheet, not just the first block, so a
#     # second batch of stones below a summary block doesn't get mistaken
#     # for outliers.
#     median_carat = None
#     if carat_col and carat_col in df.columns:
#         if stock_col and stock_col in df.columns:
#             has_stock = df[stock_col].notna() & (df[stock_col].astype(str).str.strip() != "")
#             sample_carats = pd.to_numeric(df.loc[has_stock, carat_col], errors="coerce").dropna()
#         else:
#             sample_carats = pd.to_numeric(df[carat_col], errors="coerce").dropna()
#         if not sample_carats.empty:
#             median_carat = sample_carats.median()

#     keep_mask = []
#     for _, row in df.iterrows():
#         junk = False

#         # ── condition 1: empty stock + non-empty carat ──
#         if stock_col and stock_col in df.columns and carat_col and carat_col in df.columns:
#             stock_val = row.get(stock_col)
#             carat_val = row.get(carat_col)
#             stock_empty = pd.isna(stock_val) or str(stock_val).strip() == ""
#             carat_numeric = pd.notna(carat_val) and str(carat_val).strip() != ""
#             if stock_empty and carat_numeric:
#                 junk = True

#         # ── condition 2: carat value is a huge outlier (SUM) ──
#         if not junk and carat_col and carat_col in df.columns and median_carat:
#             carat_val = pd.to_numeric(row.get(carat_col), errors="coerce")
#             if pd.notna(carat_val) and carat_val >= median_carat * 3:
#                 junk = True

#         # ── condition 3: stock ID is a summary keyword ──
#         if not junk and stock_col and stock_col in df.columns:
#             stock_val = str(row.get(stock_col, "")).strip().lower()
#             if stock_val in summary_keywords:
#                 junk = True

#         keep_mask.append(not junk)

#     return df.loc[keep_mask].reset_index(drop=True)


# def safe_filename(raw_name: str, n_stones: int) -> str:
#     """Party_Name_No_of_Stones_Current_Date.xlsx"""
#     stem = os.path.splitext(raw_name)[0]
#     stem = re.sub(r'[^\w\s\-]', '', stem).strip()
#     stem = re.sub(r'\s+', '_', stem)
#     date_str = datetime.now().strftime("%d-%m-%Y")
#     return f"{stem}_{n_stones}_Stones_{date_str}.xlsx"


# def smart_default(key: str, cols: list) -> int:
#     hints = {
#         "stock_no":    ["stock", "pexkt", "packet", "pkt", "id", "lot"],
#         "certi_no":    ["certi", "cert", "gia", "igi", "report"],
#         "shape":       ["shape"],
#         "carat":       ["carat", "ct", "weight"],
#         "lab":         ["lab", "igi", "gia"],
#         "color":       ["color", "colour"],
#         "clarity":     ["clarity", "clar"],
#         "final":       ["final", "rap%", "disc%", "back%"],
#         "vdb":         ["vdb"],
#         "usa":         ["usa"],
#     }
#     keywords = hints.get(key, [])
#     lower_cols = [c.lower() for c in cols[1:]]
#     for kw in keywords:
#         for i, lc in enumerate(lower_cols):
#             if kw in lc:
#                 return i + 1
#     return 0


# def px(hex_color: str) -> PatternFill:
#     return PatternFill("solid", fgColor=hex_color)


# def thin_border() -> Border:
#     s = Side(style="thin", color="D0D0D0")
#     return Border(left=s, right=s, top=s, bottom=s)


# def add_calculated_fields(df: pd.DataFrame, tariff_value: float, rate_value: float) -> pd.DataFrame:
#     """
#     Mirrors the exact Excel formulas (N, O, P..Z) in pandas so we can:
#       1) sort the data rows smallest → highest by column U before writing
#          them into the sheet (same order the U formula would produce), and
#       2) build the Summary sheet's insights/charts from the same numbers.
#     This does NOT change any Excel formula written by build_output_excel —
#     it's a read-only mirror used purely for ordering + reporting.
#     """
#     d = df.copy()
#     d["final_amt"]  = d["carat"] * d["final"]                 # I
#     d["vdb_amt"]    = d["vdb"] * d["carat"]                   # K
#     d["usa_amt"]    = d["usa"] * d["carat"]                   # M
#     d["tariff"]     = d["final"] * tariff_value                # N
#     d["tariff_amt"] = d["tariff"] * d["carat"]                 # O

#     sum_final_amt = d["final_amt"].sum()
#     sumproduct_ld = (d["usa"] * d["carat"]).sum()

#     d["U"] = (d["usa"] * sum_final_amt / sumproduct_ld) if sumproduct_ld else 0.0
#     d["V"] = d["U"] * d["carat"]
#     d["W"] = d["usa"] - d["U"]
#     d["X"] = (d["W"] * 100) / d["U"].replace(0, pd.NA)
#     d["Y"] = d["U"] * rate_value
#     d["total_inr"] = d["Y"] * d["carat"]                       # Z — Total INR Amt
#     return d


# # ─────────────────────────────────────────────────────────────────────────────
# #  CORE EXCEL BUILDER
# #
# #  Layout (matches Image 2 / source file exactly):
# #   Row 1        : Headers
# #   Rows 2..DS+n-1 : Data  (n = number of stones)
# #   SUM_ROW      : TOTAL
# #   AVG_ROW      : AVG / CARAT
# #   (blank rows)
# #   FINAL_ROW    : FINAL label + value
# #   TARIFF_ROW   : TARIFF x  label + value (yellow, red text) ← used by col-N formula
# #   RATE_ROW     : $ RATE label + value (yellow, red text) ← used by col-Y formula
# #   DIFF_ROW     : Difference
# #
# #  Column formulas (H onward) replicate source exactly, adjusted for output row refs:
# #   N  =ROUND(H*$N$TARIFF_ROW, 2)                 ← With Tariff = Final × Tariff multiplier
# #   O  =ROUND(N*D, 2)
# #   P  =ROUND(((L-(M_AVG - O_AVG))/1.08), 2)      ← where M_AVG=M$AVG_ROW, O_AVG=O$AVG_ROW
# #   Q  =ROUND(P*D, 2)
# #   R  =ROUND(P*1.08, 2)
# #   S  =ROUND(L-R, 2)
# #   T  =ROUND((S*100)/P, 2)
# #   U  =ROUND((L*$I$SUM_ROW)/SUMPRODUCT($L$DS:$L$DE,$D$DS:$D$DE), 2)
# #   V  =ROUND(U*D, 2)
# #   W  =ROUND(L-U, 2)
# #   X  =ROUND((W*100)/U, 2)
# #   Y  =ROUND(U*$N$RATE_ROW, 2)   ← references RATE cell instead of hardcoded 95.95
# #
# #  SUM row  : D, I, K, M, O, Q, V
# #  AVG row  : I, K, M, O, Q, V  (each SUM/D_SUM)
# # ─────────────────────────────────────────────────────────────────────────────
# def build_output_excel(mapped_df: pd.DataFrame, rate_value: float = 95.95,
#                         tariff_value: float = 1.08) -> bytes:
#     wb = openpyxl.Workbook()
#     ws = wb.active
#     ws.title = "Diamond Data"

#     n        = len(mapped_df)
#     HDR      = 1
#     DS       = 2            # data start row
#     DE       = DS + n - 1   # data end row
#     SUM_ROW    = DE + 1
#     AVG_ROW    = DE + 2
#     FINAL_ROW  = DE + 5
#     TARIFF_ROW = DE + 6
#     RATE_ROW   = DE + 7
#     DIFF_ROW   = DE + 8

#     # ── Header row ────────────────────────────────────────────────────────
#     for c, hdr in enumerate(OUTPUT_HEADERS, 1):
#         cell = ws.cell(row=HDR, column=c, value=hdr)
#         cell.fill      = px(CLR_HEADER)
#         cell.font      = Font(name="Arial", bold=True, color="FFFFFF", size=10)
#         cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
#         cell.border    = thin_border()
#     ws.row_dimensions[HDR].height = 36

#     # ── Data rows ─────────────────────────────────────────────────────────
#     for i, (_, row) in enumerate(mapped_df.iterrows()):
#         r = DS + i

#         # Raw input columns (A-M in output = cols 1-13, from party file)
#         raw = {
#             1:  row.get("stock_no"),
#             2:  row.get("certi_no"),
#             3:  row.get("shape"),
#             4:  row.get("carat"),
#             5:  row.get("lab"),
#             6:  row.get("color"),
#             7:  row.get("clarity"),
#             8:  row.get("final"),       # H  Final (%)
#             10: row.get("vdb"),         # J  VDB (%)
#             12: row.get("usa"),         # L  USA (%)
#         }

#         # Formula columns — col N (14) is now CALCULATED as Final × Tariff
#         # multiplier instead of being read from the party file.
#         formulas = {
#             9:  f"=ROUND(D{r}*H{r},2)",           # I  Final Amt
#             11: f"=ROUND(J{r}*D{r},2)",            # K  VDB Amt
#             13: f"=ROUND(L{r}*D{r},2)",            # M  USA Amt

#             # N  With Tariff = Final × Tariff multiplier (user-entered, cell ref)
#             14: f"=ROUND(H{r}*$N${TARIFF_ROW},2)",
#             # O  Tariff Amt = N*D
#             15: f"=ROUND(N{r}*D{r},2)",

#             # ── YELLOW columns (P-Y) — exact source formulas ──
#             # P: =ROUND(((L-(M_AVG-O_AVG))/1.08),2)
#             16: f"=ROUND(((L{r}-(M${AVG_ROW}-O${AVG_ROW}))/1.08),2)",
#             # Q: =ROUND(P*D,2)
#             17: f"=ROUND(P{r}*D{r},2)",
#             # R: =ROUND(P*1.08,2)
#             18: f"=ROUND(P{r}*1.08,2)",
#             # S: =ROUND(L-R,2)
#             19: f"=ROUND(L{r}-R{r},2)",
#             # T: =ROUND((S*100)/P,2)
#             20: f"=ROUND((S{r}*100)/P{r},2)",
#             # U: =ROUND((L*$I$SUM)/SUMPRODUCT($L$DS:$L$DE,$D$DS:$D$DE),2)
#             21: f"=ROUND((L{r}*$I${SUM_ROW})/SUMPRODUCT($L${DS}:$L${DE},$D${DS}:$D${DE}),2)",
#             # V: =ROUND(U*D,2)
#             22: f"=ROUND(U{r}*D{r},2)",
#             # W: =ROUND(L-U,2)
#             23: f"=ROUND(L{r}-U{r},2)",
#             # X: =ROUND((W*100)/U,2)
#             24: f"=ROUND((W{r}*100)/U{r},2)",
#             # Y: =ROUND(U*RATE,2)  — references the RATE cell
#             25: f"=ROUND(U{r}*$N${RATE_ROW},2)",
#             # Z: =ROUND(Y*D,2)  — Total INR Amt
#             26: f"=ROUND(Y{r}*D{r},2)",
#         }

#         for c in range(1, 27):
#             val = raw.get(c, formulas.get(c))
#             cell = ws.cell(row=r, column=c, value=val)
#             cell.border    = thin_border()
#             cell.alignment = Alignment(horizontal="center", vertical="center")
#             cell.font      = Font(name="Arial", size=10)
#             if c in COL_COLORS:
#                 cell.fill = px(COL_COLORS[c])
#             if c in (4, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26):
#                 cell.number_format = "0.00"

#         ws.row_dimensions[r].height = 18

#     # ── TOTAL row ─────────────────────────────────────────────────────────
#     total_fill = px("F2F2F2")
#     lbl = ws.cell(row=SUM_ROW, column=1, value="TOTAL")
#     lbl.font = Font(name="Arial", bold=True, size=10)
#     lbl.fill = total_fill; lbl.border = thin_border()
#     lbl.alignment = Alignment(horizontal="center")

#     # SUM cols: D(4), I(9), K(11), M(13), O(15), Q(17), V(22), Z(26)
#     sum_cols = {4:"D", 9:"I", 11:"K", 13:"M", 15:"O", 17:"Q", 22:"V", 26:"Z"}
#     for col, letter in sum_cols.items():
#         cell = ws.cell(row=SUM_ROW, column=col,
#                        value=f"=SUM({letter}{DS}:{letter}{DE})")
#         cell.font = Font(name="Arial", bold=True, size=10)
#         cell.fill = total_fill; cell.border = thin_border()
#         cell.alignment = Alignment(horizontal="center")
#         cell.number_format = "0.00"

#     # ── AVG / CARAT row ───────────────────────────────────────────────────
#     avg_fill = px("E2EFDA")
#     lbl2 = ws.cell(row=AVG_ROW, column=1, value="AVG / CARAT")
#     lbl2.font = Font(name="Arial", bold=True, size=10)
#     lbl2.fill = avg_fill; lbl2.border = thin_border()
#     lbl2.alignment = Alignment(horizontal="center")

#     # AVG cols: I(9), K(11), M(13), O(15), Q(17), V(22)
#     avg_cols = {9:"I", 11:"K", 13:"M", 15:"O", 17:"Q", 22:"V"}
#     for col, letter in avg_cols.items():
#         cell = ws.cell(row=AVG_ROW, column=col,
#                        value=f"=ROUND({letter}{SUM_ROW}/D{SUM_ROW},2)")
#         cell.font = Font(name="Arial", bold=True, size=10)
#         cell.fill = avg_fill; cell.border = thin_border()
#         cell.alignment = Alignment(horizontal="center")
#         cell.number_format = "0.00"

#     # ── FINAL / TARIFF / RATE / Difference label block ───────────────────
#     # FINAL row
#     for c, val in ((13, "FINAL"), (14, f"=ROUND(I{AVG_ROW},2)")):
#         cell = ws.cell(row=FINAL_ROW, column=c, value=val)
#         cell.font = Font(name="Arial", bold=True, size=10)
#         cell.border = thin_border()
#         cell.alignment = Alignment(horizontal="center")
#         cell.number_format = "0.00"

#     # TARIFF row — label M, value N (yellow fill, red bold text for value)
#     # This is the multiplier used to CALCULATE "With Tariff" (col N) as
#     # Final × Tariff, e.g. 1.08 or 1.06 — editable directly in Excel.
#     lbl_tariff = ws.cell(row=TARIFF_ROW, column=13, value="TARIFF x")
#     lbl_tariff.font = Font(name="Arial", bold=True, size=10)
#     lbl_tariff.border = thin_border()
#     lbl_tariff.alignment = Alignment(horizontal="center")

#     tariff_cell = ws.cell(row=TARIFF_ROW, column=14, value=tariff_value)
#     tariff_cell.fill   = px("FFFF00")
#     tariff_cell.font   = Font(name="Arial", bold=True, color="FF0000", size=10)
#     tariff_cell.border = thin_border()
#     tariff_cell.alignment = Alignment(horizontal="center")
#     tariff_cell.number_format = "0.00"

#     # RATE row — label M, value N (yellow fill, red bold text for value)
#     lbl_rate = ws.cell(row=RATE_ROW, column=13, value="$ RATE")
#     lbl_rate.font = Font(name="Arial", bold=True, size=10)
#     lbl_rate.border = thin_border()
#     lbl_rate.alignment = Alignment(horizontal="center")

#     rate_cell = ws.cell(row=RATE_ROW, column=14, value=rate_value)
#     rate_cell.fill   = px("FFFF00")
#     rate_cell.font   = Font(name="Arial", bold=True, color="FF0000", size=10)
#     rate_cell.border = thin_border()
#     rate_cell.alignment = Alignment(horizontal="center")
#     rate_cell.number_format = "0.00"

#     # Difference row
#     for c, val in ((13, "Difference"), (14, f"=M{AVG_ROW}-O{AVG_ROW}")):
#         cell = ws.cell(row=DIFF_ROW, column=c, value=val)
#         cell.font = Font(name="Arial", bold=True, size=10)
#         cell.border = thin_border()
#         cell.alignment = Alignment(horizontal="center")
#         cell.number_format = "0.00"

#     # ── Column widths ──────────────────────────────────────────────────────
#     widths = [14, 14, 13, 7, 6, 7, 9,  8, 11,  8, 11,  8, 11, 12, 11,
#               11, 11, 11,  8,  9, 11, 11, 11,  9, 13, 15]
#     for i, w in enumerate(widths, 1):
#         ws.column_dimensions[get_column_letter(i)].width = w

#     ws.freeze_panes = "A2"

#     # ── AutoFilter across the header + data table (A:Z) ─────────────────────
#     ws.auto_filter.ref = f"A{HDR}:Z{DE}"

#     buf = io.BytesIO()
#     wb.save(buf)
#     buf.seek(0)
#     return buf.getvalue()


# # ─────────────────────────────────────────────────────────────────────────────
# #  HERO BANNER
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="hero">
#   <div class="hero-content">
#     <p class="hero-company">✦ BRIGHT TRADING INDIA LLP</p>
#     <p class="hero-title">
#       <span class="hero-gem-wrap">
#         <span class="gem">💎</span>
#         <span class="spark spark-1">✦</span>
#         <span class="spark spark-2">✦</span>
#       </span>
#       <span class="txt">Diamond Cost Optimizer</span>
#     </p>
#     <p class="hero-sub">Smart Cost Adjustment &amp; Equal Profit Calculator</p>
#   </div>
# </div>
# """, unsafe_allow_html=True)

# st.markdown("""
# <div class="legend-wrap">
#   <div class="pill"><div class="dot" style="background:#98F6EB"></div>Carat</div>
#   <div class="pill"><div class="dot" style="background:#FFCF37"></div>Final / Amt</div>
#   <div class="pill"><div class="dot" style="background:#5991D5"></div>VDB / Amt</div>
#   <div class="pill"><div class="dot" style="background:#F79B4F"></div>USA / Amt</div>
#   <div class="pill"><div class="dot" style="background:#9E89B9"></div>With Tariff / Amt</div>
#   <div class="pill"><div class="dot" style="background:#FFFF00"></div>Calculated Columns (P–Y)</div>
# </div>
# """, unsafe_allow_html=True)

# # ─────────────────────────────────────────────────────────────────────────────
# #  STEP 1 – UPLOAD
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="step-card">
#   <div class="step-header">
#     <span class="step-num">1</span>
#     <span class="step-title">Upload Party File</span>
#   </div>
#   <p style="color:#64748b;font-size:.87rem;margin:0 0 .9rem 42px;">
#     Accepts <b style="color:#a78bfa">CSV, XLSX, XLS, XLSM, TSV</b> — any column order, any layout.
#   </p>
# </div>
# """, unsafe_allow_html=True)

# party_file = st.file_uploader(
#     "Drop your file here or click Upload",
#     type=["csv", "xlsx", "xls", "xlsm", "tsv"],
#     key="party_file",
#     label_visibility="collapsed",
# )

# if party_file is None:
#     st.markdown('<div class="banner banner-info">👆 Upload the party file above to begin.</div>',
#                 unsafe_allow_html=True)
#     st.stop()

# raw_df = read_uploaded_file(party_file)
# if raw_df is None:
#     st.stop()

# h_row = detect_header_row(raw_df)
# df = raw_df.copy()
# df.columns = raw_df.iloc[h_row]
# df = df.iloc[h_row + 1:].reset_index(drop=True)
# df.columns = dedupe_columns([str(c).strip() if pd.notna(c) else f"Col_{i}"
#                               for i, c in enumerate(df.columns)])
# df = df.dropna(how="all")

# # ── Auto-detect and strip SUM/footer rows at the bottom ──────────────────────
# # We do a quick keyword-based column sniff so trim works before Step-2 mapping
# def _sniff_col(df, keywords):
#     for col in df.columns:
#         if any(kw in str(col).lower() for kw in keywords):
#             return col
#     return None

# _stock_col = _sniff_col(df, ["stock", "packet", "pkt", "lot", "id"])
# _carat_col = _sniff_col(df, ["carat", "ct", "weight"])
# df = trim_data_rows(df, _stock_col, _carat_col)

# st.markdown(f"""
# <div class="stats-row">
#   <div class="stat-box"><div class="stat-val">{len(df)}</div><div class="stat-label">Rows Detected</div></div>
#   <div class="stat-box"><div class="stat-val">{len(df.columns)}</div><div class="stat-label">Columns Found</div></div>
#   <div class="stat-box"><div class="stat-val">{party_file.name.rsplit(".",1)[-1].upper()}</div><div class="stat-label">File Format</div></div>
# </div>
# """, unsafe_allow_html=True)

# with st.expander("👀 Preview raw data (first 5 rows)"):
#     st.dataframe(df.head(5), use_container_width=True)

# # ─────────────────────────────────────────────────────────────────────────────
# #  STEP 2 – COLUMN MAPPING
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="step-card">
#   <div class="step-header">
#     <span class="step-num">2</span>
#     <span class="step-title">Map Columns</span>
#   </div>
# </div>
# """, unsafe_allow_html=True)
# st.caption("Select which column in the party file corresponds to each required field.")

# available = ["— skip / not available —"] + list(df.columns)
# mapping   = {}

# cols_left, cols_right = st.columns(2)
# for idx, (key, label) in enumerate(REQUIRED_COLS):
#     default = smart_default(key, available)
#     container = cols_left if idx % 2 == 0 else cols_right
#     with container:
#         choice = st.selectbox(f"**{label}**", options=available,
#                               index=default, key=f"map_{key}")
#         mapping[key] = None if choice.startswith("—") else choice

# # ── Tariff multiplier + $ RATE inputs ─────────────────────────────────────────
# st.markdown("---")
# st.markdown("**Tariff Multiplier** — used to calculate the `With Tariff` column (`Final × Multiplier`), e.g. 1.08 or 1.06")
# tariff_input = st.number_input("Tariff multiplier", min_value=1.0, value=1.08, step=0.01, format="%.2f")

# st.markdown("**$ Rate** — used in the final Y-column formula (`U × $ Rate`)")
# rate_input = st.number_input("$ Rate value", min_value=0.0, value=95.95, step=0.01, format="%.2f")

# mandatory = ["carat", "final", "vdb", "usa"]
# missing   = [next(l for kk, l in REQUIRED_COLS if kk == k)
#              for k in mandatory if not mapping.get(k)]
# if missing:
#     st.markdown(f'<div class="banner banner-warn">⚠️ Please map required columns: <b>{", ".join(missing)}</b></div>',
#                 unsafe_allow_html=True)
#     st.stop()

# # ─────────────────────────────────────────────────────────────────────────────
# #  STEP 3 – PREVIEW & DOWNLOAD
# # ─────────────────────────────────────────────────────────────────────────────
# st.markdown("""
# <div class="step-card">
#   <div class="step-header">
#     <span class="step-num">3</span>
#     <span class="step-title">Preview &amp; Download</span>
#   </div>
# </div>
# """, unsafe_allow_html=True)

# # ── Re-trim using the user's actual mapped columns (catches any edge cases) ──
# df = trim_data_rows(df, mapping.get("stock_no"), mapping.get("carat"))

# records = []
# for _, row in df.iterrows():
#     rec = {key: (row[col] if col else None) for key, col in mapping.items()}
#     records.append(rec)

# mapped_df = pd.DataFrame(records)
# for col in ["carat", "final", "vdb", "usa"]:
#     mapped_df[col] = pd.to_numeric(mapped_df[col], errors="coerce")

# mapped_df = mapped_df.dropna(subset=["carat", "final", "vdb", "usa"], how="all")
# n_stones  = len(mapped_df)

# if n_stones == 0:
#     st.markdown('<div class="banner banner-warn">⚠️ No valid rows found. Check your column mappings.</div>',
#                 unsafe_allow_html=True)
#     st.stop()

# # ── Arrange rows smallest → highest by column U (same order the U formula
# #    itself produces in Excel) ──────────────────────────────────────────────
# _sort_calc = add_calculated_fields(mapped_df, tariff_input, rate_input)
# mapped_df = (mapped_df.assign(_u=_sort_calc["U"].values)
#                        .sort_values("_u", kind="mergesort")
#                        .drop(columns="_u")
#                        .reset_index(drop=True))

# st.markdown(f'<div class="banner banner-success">✅ <b>{n_stones} stones</b> ready to export.</div>',
#             unsafe_allow_html=True)

# preview_rename = {
#     "stock_no":"Stock ID.", "certi_no":"Certi No.", "shape":"Shape",
#     "carat":"Carat", "lab":"Lab", "color":"Color", "clarity":"Clarity",
#     "final":"Final", "vdb":"VDB", "usa":"USA",
# }
# st.dataframe(mapped_df.rename(columns=preview_rename).head(10), use_container_width=True)

# with st.spinner("Building Excel with formulas and colours…"):
#     xlsx_bytes   = build_output_excel(mapped_df, rate_value=rate_input, tariff_value=tariff_input)
#     out_filename = safe_filename(party_file.name, n_stones)

# st.markdown('<div class="banner banner-success">💾 File ready — all formulas match the reference format exactly (cols P–Y yellow).</div>',
#             unsafe_allow_html=True)

# st.download_button(
#     label=f"⬇️  Download  {out_filename}",
#     data=xlsx_bytes,
#     file_name=out_filename,
#     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#     use_container_width=True,
# )

# with st.expander("📐 Formula & Colour Reference"):
#     st.markdown(f"""
# | Col | Header | Colour | Formula / Source |
# |---|---|---|---|
# | A | Stock ID | — | From party file |
# | B | Certi No | — | From party file |
# | C | Shape | — | From party file |
# | D | Carat | 🟦 Teal | From party file |
# | E | Lab | — | From party file |
# | F | Color | — | From party file |
# | G | Clarity | — | From party file |
# | H | Final | 🟨 Gold | From party file |
# | I | Amt | 🟨 Gold | `D × H` |
# | J | VDB | 🔵 Blue | From party file |
# | K | Amt | 🔵 Blue | `J × D` |
# | L | USA | 🟠 Orange | From party file |
# | M | Amt | 🟠 Orange | `L × D` |
# | N | With Tariff | 🟣 Lavender | `H × Tariff Multiplier` — Multiplier = **{tariff_input}** |
# | O | Amt | 🟣 Lavender | `N × D` |
# | P | — | 🟡 Yellow | `ROUND(((L-(M_avg−O_avg))/1.08),2)` |
# | Q | — | 🟡 Yellow | `ROUND(P×D, 2)` |
# | R | — | 🟡 Yellow | `ROUND(P×1.08, 2)` |
# | S | — | 🟡 Yellow | `ROUND(L−R, 2)` |
# | T | — | 🟡 Yellow | `ROUND((S×100)/P, 2)` |
# | U | — | 🟡 Yellow | `ROUND((L×ΣI) / SUMPRODUCT(ΣL,ΣD), 2)` |
# | V | — | 🟡 Yellow | `ROUND(U×D, 2)` |
# | W | — | 🟡 Yellow | `ROUND(L−U, 2)` |
# | X | — | 🟡 Yellow | `ROUND((W×100)/U, 2)` |
# | Y | — | 🟡 Yellow | `ROUND(U×$RATE, 2)` — RATE = **{rate_input}** |

# > 🟡 Both **TARIFF x** and **$ RATE** cells are highlighted yellow/red in the summary block — editable directly in Excel.
# > 🔽 Rows are sorted smallest → highest by column **U**, and an **AutoFilter (A:Z)** is enabled on the data table.
# """)




import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import re
import os
from datetime import datetime

st.set_page_config(page_title="💎 Diamond Formatter", page_icon="💎", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"] { background-color: #0d0f14 !important; color: #e2e8f0 !important; }
[data-testid="stHeader"] { background: #0d0f14 !important; }
[data-testid="stSidebar"] { background: #111318 !important; }
section[data-testid="stMain"] > div { background: #0d0f14 !important; }
.block-container { padding-top: 1.5rem !important; }
.hero { position:relative;background:linear-gradient(135deg,#0f1629 0%,#1a0a2e 50%,#0a1628 100%);background-size:200% 200%;animation:heroGradientShift 12s ease infinite;border:1px solid #2a2d3e;border-radius:20px;padding:2.4rem 2rem;margin-bottom:2rem;overflow:hidden;box-shadow:0 10px 40px rgba(0,0,0,.35); }
@keyframes heroGradientShift { 0%{background-position:0% 50%} 50%{background-position:100% 50%} 100%{background-position:0% 50%} }
.hero::before,.hero::after { content:'';position:absolute;border-radius:50%;filter:blur(60px);opacity:.55;animation:heroFloat 8s ease-in-out infinite;pointer-events:none; }
.hero::before { width:240px;height:240px;background:rgba(99,102,241,.35);top:-70px;left:8%; }
.hero::after { width:280px;height:280px;background:rgba(168,85,247,.28);bottom:-90px;right:8%;animation-delay:-4s; }
@keyframes heroFloat { 0%,100%{transform:translateY(0) translateX(0)} 50%{transform:translateY(-22px) translateX(16px)} }
.hero-content { position:relative;z-index:1;text-align:left; }
.hero-company { display:inline-block;font-size:.92rem;font-weight:700;letter-spacing:2.6px;color:#c4b5fd;text-transform:uppercase;padding:7px 20px;border:1px solid rgba(167,139,250,.45);border-radius:30px;background:rgba(99,102,241,.12);margin:0 0 1.1rem;animation:heroFadeInDown .8s ease both,heroPulseBorder 3s ease-in-out infinite 1s; }
@keyframes heroFadeInDown { from{opacity:0;transform:translateY(-14px)} to{opacity:1;transform:translateY(0)} }
@keyframes heroPulseBorder { 0%,100%{box-shadow:0 0 0 rgba(167,139,250,0)} 50%{box-shadow:0 0 20px rgba(167,139,250,.4)} }
.hero-title { display:flex;align-items:center;gap:12px;margin:0 0 .6rem;animation:heroFadeInUp .9s ease .2s both; }
.hero-gem-wrap { position:relative;display:inline-flex;align-items:center;justify-content:center; }
.hero-title .gem { display:inline-block;font-size:2rem;line-height:1;transform-style:preserve-3d;animation:heroGemSpin 4.5s ease-in-out infinite; }
.hero-gem-wrap .spark { position:absolute;color:#7dd3fc;font-size:.6rem;opacity:0; }
.hero-gem-wrap .spark-1 { top:-6px;right:-8px;animation:heroTwinkle 2.6s ease-in-out infinite; }
.hero-gem-wrap .spark-2 { bottom:-4px;left:-10px;animation:heroTwinkle 2.6s ease-in-out infinite 1.1s; }
@keyframes heroTwinkle { 0%,100%{opacity:0;transform:scale(.4) rotate(0deg)} 50%{opacity:1;transform:scale(1.1) rotate(25deg)} }
@keyframes heroGemSpin { 0%{transform:perspective(320px) rotateY(0deg) scale(1);filter:drop-shadow(0 0 6px rgba(167,139,250,.45))} 45%{transform:perspective(320px) rotateY(180deg) scale(1.1);filter:drop-shadow(0 0 16px rgba(56,189,248,.8))} 55%{transform:perspective(320px) rotateY(180deg) scale(1.1);filter:drop-shadow(0 0 16px rgba(56,189,248,.8))} 100%{transform:perspective(320px) rotateY(360deg) scale(1);filter:drop-shadow(0 0 6px rgba(167,139,250,.45))} }
.hero-title .txt { font-size:2.3rem;font-weight:800;letter-spacing:-.5px;line-height:1.15;background:linear-gradient(90deg,#a78bfa,#6366f1,#38bdf8,#a78bfa);background-size:300% auto;-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;animation:heroShimmer 6s linear infinite; }
@keyframes heroShimmer { to{background-position:300% center} }
@keyframes heroFadeInUp { from{opacity:0;transform:translateY(18px)} to{opacity:1;transform:translateY(0)} }
.hero-sub { color:#94a3b8;font-size:1rem;margin:0;animation:heroFadeInUp .9s ease .45s both; }
.legend-wrap { display:flex;flex-wrap:wrap;gap:8px;margin:1rem 0 1.5rem; }
.pill { display:inline-flex;align-items:center;gap:7px;background:#1a1d27;border:1px solid #2a2d3e;border-radius:20px;padding:5px 12px;font-size:.78rem;color:#cbd5e1;transition:transform .18s ease,border-color .18s ease; }
.pill:hover { transform:translateY(-2px);border-color:#6366f1; }
.dot { width:12px;height:12px;border-radius:3px;flex-shrink:0; }
.step-card { background:#13161f;border:1px solid #1e2130;border-radius:12px;padding:1.5rem;margin-bottom:1rem; }
.step-num { display:inline-flex;align-items:center;justify-content:center;width:32px;height:32px;border-radius:50%;background:linear-gradient(135deg,#6366f1,#8b5cf6);color:white;font-weight:700;font-size:.9rem;margin-right:10px;flex-shrink:0; }
.step-title { font-size:1.1rem;font-weight:600;color:#f1f5f9; }
.step-header { display:flex;align-items:center;margin-bottom:1rem; }
[data-testid="stFileUploader"] { background:#13161f !important;border:2px dashed #2a2d3e !important;border-radius:10px !important; }
[data-testid="stFileUploader"]:hover { border-color:#6366f1 !important; }
.stSelectbox > div > div { background:#1a1d27 !important;border:1px solid #2a2d3e !important;color:#e2e8f0 !important;border-radius:8px !important; }
.stSelectbox label { color:#cbd5e1 !important;font-size:.85rem !important; }
.stDownloadButton button { background:linear-gradient(135deg,#6366f1,#8b5cf6) !important;color:white !important;border:none !important;border-radius:10px !important;font-weight:600 !important;padding:.75rem 1.5rem !important;font-size:1rem !important;transition:all .2s !important;box-shadow:0 4px 20px rgba(99,102,241,.4) !important; }
.stDownloadButton button:hover { transform:translateY(-2px) !important;box-shadow:0 8px 30px rgba(99,102,241,.6) !important; }
.banner { border-radius:10px;padding:.85rem 1.1rem;font-size:.9rem;margin:.5rem 0;display:flex;align-items:center;gap:10px; }
.banner-success { background:rgba(34,197,94,.1);border:1px solid rgba(34,197,94,.3);color:#86efac; }
.banner-info    { background:rgba(99,102,241,.1);border:1px solid rgba(99,102,241,.3);color:#a5b4fc; }
.banner-warn    { background:rgba(234,179,8,.1);border:1px solid rgba(234,179,8,.3);color:#fde68a; }
.stats-row { display:flex;gap:12px;flex-wrap:wrap;margin:1rem 0; }
.stat-box { background:#13161f;border:1px solid #1e2130;border-radius:10px;padding:.75rem 1.25rem;flex:1;min-width:130px; }
.stat-val { font-size:1.6rem;font-weight:700;color:#a78bfa; }
.stat-label { font-size:.75rem;color:#64748b;text-transform:uppercase;letter-spacing:.5px; }
hr { border-color:#1e2130 !important;margin:1.5rem 0 !important; }
[data-testid="stExpander"] { background:#13161f !important;border:1px solid #1e2130 !important;border-radius:10px !important; }
summary { color:#94a3b8 !important; }
</style>
""", unsafe_allow_html=True)

# ── Colors ────────────────────────────────────────────────────────────────────
CLR_HEADER = "1F4E79"
CLR_CARAT  = "98F6EB"
CLR_FINAL  = "FFCF37"
CLR_VDB    = "5991D5"
CLR_USA    = "F79B4F"
CLR_TARIFF = "9E89B9"
CLR_YELLOW = "FFFF00"

COL_COLORS = {4:CLR_CARAT, 8:CLR_FINAL, 9:CLR_FINAL, 10:CLR_VDB, 11:CLR_VDB,
              12:CLR_USA, 13:CLR_USA, 14:CLR_TARIFF, 15:CLR_TARIFF}
for _c in range(16, 27):
    COL_COLORS[_c] = CLR_YELLOW

OUTPUT_HEADERS = [
    "Stock ID.","Certi No.","Shape","Carat","Lab","Color","Clarity",
    "Final","Amt","VDB","Amt","USA","Amt","With Tariff","Amt",
    "","","","","","","","","","","Total INR Amt",
]

REQUIRED_COLS = [
    ("stock_no", "Stock ID."), ("certi_no", "Certi No."), ("shape",   "Shape"),
    ("carat",    "Carat"),     ("lab",      "Lab"),        ("color",   "Color"),
    ("clarity",  "Clarity"),   ("final",    "Final"),      ("vdb",     "VDB"),
    ("usa",      "USA"),
]

# ── Helpers ───────────────────────────────────────────────────────────────────
def read_uploaded_file(f):
    name = f.name.lower()
    try:
        if name.endswith(".csv"):   return pd.read_csv(f)
        if name.endswith((".xlsx",".xls",".xlsm")): return pd.read_excel(f, header=None)
        if name.endswith(".tsv"):   return pd.read_csv(f, sep="\t")
        st.error("Unsupported format."); return None
    except Exception as e:
        st.error(f"Could not read file: {e}"); return None


def dedupe_columns(cols):
    seen, out = {}, []
    for c in cols:
        name = str(c).strip() if c is not None and str(c).strip() else "Col"
        if name not in seen:
            seen[name] = 0; out.append(name)
        else:
            seen[name] += 1; out.append(f"{name}_{seen[name]}")
    return out


def detect_header_row(df: pd.DataFrame) -> int:
    for i, row in df.iterrows():
        vals = [str(v).strip() for v in row if pd.notna(v) and str(v).strip()]
        if len(vals) >= 4 and not all(re.sub(r'[.\-]', '', v).isdigit() for v in vals):
            return i
    return 0


def trim_data_rows(df: pd.DataFrame, stock_col, carat_col) -> pd.DataFrame:
    """
    Stop reading rows when we hit a SUM/footer row. Three conditions:

    1. stock-ID cell is EMPTY but carat cell has a value
       → classic SUM row pattern (stock blank, carat = total)

    2. carat value exceeds an ABSOLUTE cap of 20 ct
       → no single diamond in a normal parcel exceeds 20 ct;
         a value like 28, 99, 131 is always a SUM total.
       (Previously used 3× median which wrongly cut at 6.54 ct.)

    3. stock-ID cell contains a known summary keyword
       (TOTAL, AVG, SUM, RATE, FINAL, Difference, etc.)
    """
    if df.empty:
        return df

    CARAT_CAP = 20.0          # absolute max single-stone carat — raise if needed
    SUMMARY_KW = {"total", "avg", "average", "sum", "rate", "final",
                  "difference", "subtotal", "done", "fianl", "tariff x",
                  "$ rate", "avg / carat"}

    cut_at = len(df)

    for i, (_, row) in enumerate(df.iterrows()):
        # ── condition 1: empty stock + non-empty carat ────────────────────
        if stock_col and stock_col in df.columns and carat_col and carat_col in df.columns:
            sv = row.get(stock_col)
            cv = row.get(carat_col)
            if (pd.isna(sv) or str(sv).strip() == "") and \
               (pd.notna(cv) and str(cv).strip() != ""):
                cut_at = i; break

        # ── condition 2: carat exceeds absolute cap ───────────────────────
        if carat_col and carat_col in df.columns:
            cv = pd.to_numeric(row.get(carat_col), errors="coerce")
            if pd.notna(cv) and cv > CARAT_CAP:
                cut_at = i; break

        # ── condition 3: stock cell is a summary keyword ──────────────────
        if stock_col and stock_col in df.columns:
            sv = str(row.get(stock_col, "")).strip().lower()
            if sv in SUMMARY_KW:
                cut_at = i; break

    return df.iloc[:cut_at].reset_index(drop=True)


def safe_filename(raw_name: str, n_stones: int) -> str:
    stem = re.sub(r'[^\w\s\-]', '', os.path.splitext(raw_name)[0]).strip()
    stem = re.sub(r'\s+', '_', stem)
    return f"{stem}_{n_stones}_Stones_{datetime.now().strftime('%d-%m-%Y')}.xlsx"


def smart_default(key: str, cols: list) -> int:
    hints = {
        # ⚠️  "id" removed — too generic, matches 'Grid', 'Available', etc.
        "stock_no": ["stock", "pexkt", "packet", "pkt", "refno", "ref no", "lot"],
        "certi_no": ["certi", "cert", "gia", "igi", "report", "reportno"],
        "shape":    ["shape"],
        "carat":    ["carat", " cts", "^cts$", "weight", " ct"],
        "lab":      ["^lab$", " lab", "lab "],
        "color":    ["color", "colour"],
        "clarity":  ["clarity", "clar"],
        "final":    ["final", "ma price", "rap%", "disc%", "back%"],
        "vdb":      ["vdb"],
        "usa":      ["vdb usa", "usa"],
    }
    kws = hints.get(key, [])
    lower_cols = [c.lower() for c in cols[1:]]
    for kw in kws:
        for i, lc in enumerate(lower_cols):
            # strip regex anchors for simple substring match
            kw_plain = kw.lstrip("^").rstrip("$")
            if kw_plain in lc:
                return i + 1
    return 0


def px(h): return PatternFill("solid", fgColor=h)
def thin_border():
    s = Side(style="thin", color="D0D0D0")
    return Border(left=s, right=s, top=s, bottom=s)


def build_output_excel(mapped_df: pd.DataFrame,
                       rate_value: float = 95.95,
                       tariff_value: float = 1.08) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Diamond Data"

    n = len(mapped_df)
    HDR = 1; DS = 2; DE = DS + n - 1
    SUM_ROW   = DE + 1; AVG_ROW   = DE + 2
    FINAL_ROW = DE + 5; TARIFF_ROW = DE + 6
    RATE_ROW  = DE + 7; DIFF_ROW  = DE + 8

    # Header
    for c, hdr in enumerate(OUTPUT_HEADERS, 1):
        cell = ws.cell(row=HDR, column=c, value=hdr)
        cell.fill = px(CLR_HEADER)
        cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border()
    ws.row_dimensions[HDR].height = 36

    # Data rows
    for i, (_, row) in enumerate(mapped_df.iterrows()):
        r = DS + i
        raw = {
            1: row.get("stock_no"), 2: row.get("certi_no"), 3: row.get("shape"),
            4: row.get("carat"),    5: row.get("lab"),       6: row.get("color"),
            7: row.get("clarity"),  8: row.get("final"),
            10: row.get("vdb"),     12: row.get("usa"),
        }
        formulas = {
            9:  f"=ROUND(D{r}*H{r},2)",
            11: f"=ROUND(J{r}*D{r},2)",
            13: f"=ROUND(L{r}*D{r},2)",
            14: f"=ROUND(H{r}*$N${TARIFF_ROW},2)",
            15: f"=ROUND(N{r}*D{r},2)",
            16: f"=ROUND(((L{r}-(M${AVG_ROW}-O${AVG_ROW}))/1.08),2)",
            17: f"=ROUND(P{r}*D{r},2)",
            18: f"=ROUND(P{r}*1.08,2)",
            19: f"=ROUND(L{r}-R{r},2)",
            20: f"=ROUND((S{r}*100)/P{r},2)",
            21: f"=ROUND((L{r}*$I${SUM_ROW})/SUMPRODUCT($L${DS}:$L${DE},$D${DS}:$D${DE}),2)",
            22: f"=ROUND(U{r}*D{r},2)",
            23: f"=ROUND(L{r}-U{r},2)",
            24: f"=ROUND((W{r}*100)/U{r},2)",
            25: f"=ROUND(U{r}*$N${RATE_ROW},2)",
            26: f"=ROUND(Y{r}*D{r},2)",
        }
        for c in range(1, 27):
            val = raw.get(c, formulas.get(c))
            cell = ws.cell(row=r, column=c, value=val)
            cell.border = thin_border()
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.font = Font(name="Arial", size=10)
            if c in COL_COLORS: cell.fill = px(COL_COLORS[c])
            if c in (4,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26):
                cell.number_format = "0.00"
        ws.row_dimensions[r].height = 18

    # TOTAL row
    tf = px("F2F2F2")
    lbl = ws.cell(row=SUM_ROW, column=1, value="TOTAL")
    lbl.font=Font(name="Arial",bold=True,size=10); lbl.fill=tf; lbl.border=thin_border(); lbl.alignment=Alignment(horizontal="center")
    for col,letter in {4:"D",9:"I",11:"K",13:"M",15:"O",17:"Q",22:"V",26:"Z"}.items():
        cell=ws.cell(row=SUM_ROW,column=col,value=f"=SUM({letter}{DS}:{letter}{DE})")
        cell.font=Font(name="Arial",bold=True,size=10); cell.fill=tf; cell.border=thin_border()
        cell.alignment=Alignment(horizontal="center"); cell.number_format="0.00"

    # AVG row
    af = px("E2EFDA")
    lbl2=ws.cell(row=AVG_ROW,column=1,value="AVG / CARAT")
    lbl2.font=Font(name="Arial",bold=True,size=10); lbl2.fill=af; lbl2.border=thin_border(); lbl2.alignment=Alignment(horizontal="center")
    for col,letter in {9:"I",11:"K",13:"M",15:"O",17:"Q",22:"V"}.items():
        cell=ws.cell(row=AVG_ROW,column=col,value=f"=ROUND({letter}{SUM_ROW}/D{SUM_ROW},2)")
        cell.font=Font(name="Arial",bold=True,size=10); cell.fill=af; cell.border=thin_border()
        cell.alignment=Alignment(horizontal="center"); cell.number_format="0.00"

    # FINAL / TARIFF / RATE / Difference
    for c,val in ((13,"FINAL"),(14,f"=ROUND(I{AVG_ROW},2)")):
        cell=ws.cell(row=FINAL_ROW,column=c,value=val)
        cell.font=Font(name="Arial",bold=True,size=10); cell.border=thin_border()
        cell.alignment=Alignment(horizontal="center"); cell.number_format="0.00"

    lbl_t=ws.cell(row=TARIFF_ROW,column=13,value="TARIFF x")
    lbl_t.font=Font(name="Arial",bold=True,size=10); lbl_t.border=thin_border(); lbl_t.alignment=Alignment(horizontal="center")
    tc=ws.cell(row=TARIFF_ROW,column=14,value=tariff_value)
    tc.fill=px("FFFF00"); tc.font=Font(name="Arial",bold=True,color="FF0000",size=10)
    tc.border=thin_border(); tc.alignment=Alignment(horizontal="center"); tc.number_format="0.00"

    lbl_r=ws.cell(row=RATE_ROW,column=13,value="$ RATE")
    lbl_r.font=Font(name="Arial",bold=True,size=10); lbl_r.border=thin_border(); lbl_r.alignment=Alignment(horizontal="center")
    rc=ws.cell(row=RATE_ROW,column=14,value=rate_value)
    rc.fill=px("FFFF00"); rc.font=Font(name="Arial",bold=True,color="FF0000",size=10)
    rc.border=thin_border(); rc.alignment=Alignment(horizontal="center"); rc.number_format="0.00"

    for c,val in ((13,"Difference"),(14,f"=M{AVG_ROW}-O{AVG_ROW}")):
        cell=ws.cell(row=DIFF_ROW,column=c,value=val)
        cell.font=Font(name="Arial",bold=True,size=10); cell.border=thin_border()
        cell.alignment=Alignment(horizontal="center"); cell.number_format="0.00"

    widths=[14,14,13,7,6,7,9,8,11,8,11,8,11,12,11,11,11,11,8,9,11,11,11,9,13,15]
    for i,w in enumerate(widths,1):
        ws.column_dimensions[get_column_letter(i)].width=w
    ws.freeze_panes="A2"
    ws.auto_filter.ref=f"A{HDR}:Z{DE}"

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()


# ── UI ────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero"><div class="hero-content">
  <p class="hero-company">✦ BRIGHT TRADING INDIA LLP</p>
  <p class="hero-title">
    <span class="hero-gem-wrap"><span class="gem">💎</span><span class="spark spark-1">✦</span><span class="spark spark-2">✦</span></span>
    <span class="txt">Diamond Cost Optimizer</span>
  </p>
  <p class="hero-sub">Smart Cost Adjustment &amp; Equal Profit Calculator</p>
</div></div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="legend-wrap">
  <div class="pill"><div class="dot" style="background:#98F6EB"></div>Carat</div>
  <div class="pill"><div class="dot" style="background:#FFCF37"></div>Final / Amt</div>
  <div class="pill"><div class="dot" style="background:#5991D5"></div>VDB / Amt</div>
  <div class="pill"><div class="dot" style="background:#F79B4F"></div>USA / Amt</div>
  <div class="pill"><div class="dot" style="background:#9E89B9"></div>With Tariff / Amt</div>
  <div class="pill"><div class="dot" style="background:#FFFF00"></div>Calculated Columns (P–Z)</div>
</div>
""", unsafe_allow_html=True)

# STEP 1
st.markdown("""
<div class="step-card"><div class="step-header"><span class="step-num">1</span><span class="step-title">Upload Party File</span></div>
<p style="color:#64748b;font-size:.87rem;margin:0 0 .9rem 42px;">Accepts <b style="color:#a78bfa">CSV, XLSX, XLS, XLSM, TSV</b> — any column order, any layout.</p></div>
""", unsafe_allow_html=True)

party_file = st.file_uploader("Drop your file here", type=["csv","xlsx","xls","xlsm","tsv"],
                               key="party_file", label_visibility="collapsed")

if party_file is None:
    st.markdown('<div class="banner banner-info">👆 Upload the party file above to begin.</div>', unsafe_allow_html=True)
    st.stop()

raw_df = read_uploaded_file(party_file)
if raw_df is None: st.stop()

h_row = detect_header_row(raw_df)
df = raw_df.copy()
df.columns = raw_df.iloc[h_row]
df = df.iloc[h_row + 1:].reset_index(drop=True)
df.columns = dedupe_columns(df.columns.tolist())
df = df.dropna(how="all")

# ── Sniff stock/carat cols for trim — FIXED: no 'id' keyword ─────────────────
def _sniff_col(df, keywords):
    """Exact-word sniff: keyword must be a whole word inside the column name."""
    for col in df.columns:
        col_l = str(col).lower()
        for kw in keywords:
            # whole-word match to avoid 'id' hitting 'Grid', 'Available', etc.
            if re.search(r'\b' + re.escape(kw) + r'\b', col_l):
                return col
    return None

_stock_col = _sniff_col(df, ["stock", "packet", "pkt", "refno", "ref"])
_carat_col = _sniff_col(df, ["carat", "cts", "ct", "weight"])
df = trim_data_rows(df, _stock_col, _carat_col)

st.markdown(f"""
<div class="stats-row">
  <div class="stat-box"><div class="stat-val">{len(df)}</div><div class="stat-label">Rows Detected</div></div>
  <div class="stat-box"><div class="stat-val">{len(df.columns)}</div><div class="stat-label">Columns Found</div></div>
  <div class="stat-box"><div class="stat-val">{party_file.name.rsplit(".",1)[-1].upper()}</div><div class="stat-label">File Format</div></div>
</div>
""", unsafe_allow_html=True)

with st.expander("👀 Preview raw data (first 5 rows)"):
    st.dataframe(df.head(5), use_container_width=True)

# STEP 2
st.markdown("""
<div class="step-card"><div class="step-header"><span class="step-num">2</span><span class="step-title">Map Columns</span></div></div>
""", unsafe_allow_html=True)
st.caption("Select which column corresponds to each required field. Auto-matched where possible.")

available = ["— skip / not available —"] + list(df.columns)
mapping = {}
cols_left, cols_right = st.columns(2)
for idx, (key, label) in enumerate(REQUIRED_COLS):
    default = smart_default(key, available)
    with (cols_left if idx % 2 == 0 else cols_right):
        choice = st.selectbox(f"**{label}**", options=available, index=default, key=f"map_{key}")
        mapping[key] = None if choice.startswith("—") else choice

st.markdown("---")
tariff_input = st.number_input("Tariff multiplier (used to calculate With Tariff = Final × multiplier)",
                                min_value=1.0, value=1.08, step=0.01, format="%.2f")
rate_input   = st.number_input("$ Rate value (used in Y = U × Rate)",
                                min_value=0.0, value=95.95, step=0.01, format="%.2f")

mandatory = ["carat", "final", "vdb", "usa"]
missing = [next(l for kk,l in REQUIRED_COLS if kk==k) for k in mandatory if not mapping.get(k)]
if missing:
    st.markdown(f'<div class="banner banner-warn">⚠️ Please map: <b>{", ".join(missing)}</b></div>', unsafe_allow_html=True)
    st.stop()

# STEP 3
st.markdown("""
<div class="step-card"><div class="step-header"><span class="step-num">3</span><span class="step-title">Preview &amp; Download</span></div></div>
""", unsafe_allow_html=True)

# Re-trim with confirmed mapped columns
df = trim_data_rows(df, mapping.get("stock_no"), mapping.get("carat"))

records = [{key: (row[col] if col else None) for key,col in mapping.items()} for _,row in df.iterrows()]
mapped_df = pd.DataFrame(records)
for col in ["carat","final","vdb","usa"]:
    mapped_df[col] = pd.to_numeric(mapped_df[col], errors="coerce")
mapped_df = mapped_df.dropna(subset=["carat","final","vdb","usa"], how="all")
n_stones = len(mapped_df)

if n_stones == 0:
    st.markdown('<div class="banner banner-warn">⚠️ No valid rows found. Check your column mappings.</div>', unsafe_allow_html=True)
    st.stop()

st.markdown(f'<div class="banner banner-success">✅ <b>{n_stones} stones</b> ready to export.</div>', unsafe_allow_html=True)

preview_rename = {"stock_no":"Stock ID.","certi_no":"Certi No.","shape":"Shape","carat":"Carat",
                  "lab":"Lab","color":"Color","clarity":"Clarity","final":"Final","vdb":"VDB","usa":"USA"}
st.dataframe(mapped_df.rename(columns=preview_rename).head(10), use_container_width=True)

with st.spinner("Building Excel…"):
    xlsx_bytes   = build_output_excel(mapped_df, rate_value=rate_input, tariff_value=tariff_input)
    out_filename = safe_filename(party_file.name, n_stones)

st.markdown('<div class="banner banner-success">💾 File ready.</div>', unsafe_allow_html=True)
st.download_button(label=f"⬇️  Download  {out_filename}", data=xlsx_bytes, file_name=out_filename,
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   use_container_width=True)

with st.expander("📐 Formula & Colour Reference"):
    st.markdown(f"""
| Col | Header | Colour | Formula / Source |
|---|---|---|---|
| A | Stock ID | — | Party file |
| B | Certi No | — | Party file |
| C | Shape | — | Party file |
| D | Carat | 🟦 Teal | Party file |
| E | Lab | — | Party file |
| F | Color | — | Party file |
| G | Clarity | — | Party file |
| H | Final | 🟨 Gold | Party file |
| I | Amt | 🟨 Gold | `D × H` |
| J | VDB | 🔵 Blue | Party file |
| K | Amt | 🔵 Blue | `J × D` |
| L | USA | 🟠 Orange | Party file |
| M | Amt | 🟠 Orange | `L × D` |
| N | With Tariff | 🟣 Lavender | `H × {tariff_input}` (TARIFF cell) |
| O | Amt | 🟣 Lavender | `N × D` |
| P–X | — | 🟡 Yellow | Adj cost / profit formulas |
| Y | — | 🟡 Yellow | `U × $ RATE ({rate_input})` |
| Z | Total INR Amt | 🟡 Yellow | `Y × D` |
""")