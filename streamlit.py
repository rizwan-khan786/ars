
# # # # # # # # # import streamlit as st
# # # # # # # # # import pandas as pd
# # # # # # # # # import numpy as np
# # # # # # # # # import plotly.express as px
# # # # # # # # # import re
# # # # # # # # # from datetime import datetime

# # # # # # # # # st.set_page_config(page_title="GHG Summary Analyzer", layout="wide")
# # # # # # # # # st.title("GHG Summary Analyzer")
# # # # # # # # # st.caption("Upload one or more Excel files • Improved Month-wise detection")

# # # # # # # # # # ---------------- Helpers (kept same) ----------------
# # # # # # # # # def load_excel(uploaded_file):
# # # # # # # # #     for eng in ("openpyxl", "calamine"):
# # # # # # # # #         try:
# # # # # # # # #             return pd.ExcelFile(uploaded_file, engine=eng), eng
# # # # # # # # #         except Exception:
# # # # # # # # #             continue
# # # # # # # # #     raise ImportError("No Excel engine available.")

# # # # # # # # # def find_sheet(xls):
# # # # # # # # #     if "Summary Sheet" in xls.sheet_names:
# # # # # # # # #         return "Summary Sheet"
# # # # # # # # #     for s in xls.sheet_names:
# # # # # # # # #         if "summary" in s.lower():
# # # # # # # # #             return s
# # # # # # # # #     return xls.sheet_names[0]

# # # # # # # # # def header_detect_clean(df_raw: pd.DataFrame) -> pd.DataFrame:
# # # # # # # # #     raw = df_raw.dropna(how="all").dropna(axis=1, how="all").copy()
# # # # # # # # #     header_idx = raw.head(min(15, len(raw))).notna().sum(axis=1).idxmax()
# # # # # # # # #     df = raw.copy()
# # # # # # # # #     df.columns = df.loc[header_idx].astype(str).str.strip()
# # # # # # # # #     df = df.loc[header_idx + 1 :].reset_index(drop=True)

# # # # # # # # #     seen, cols = {}, []
# # # # # # # # #     for c in df.columns:
# # # # # # # # #         base = c if c and c != "nan" else "Unnamed"
# # # # # # # # #         seen[base] = seen.get(base, 0) + 1
# # # # # # # # #         cols.append(base if seen[base] == 1 else f"{base}_{seen[base]}")
# # # # # # # # #     df.columns = cols

# # # # # # # # #     for c in df.columns:
# # # # # # # # #         parsed = pd.to_numeric(df[c], errors="coerce")
# # # # # # # # #         if parsed.notna().mean() >= 0.6:
# # # # # # # # #             df[c] = parsed
# # # # # # # # #     return df

# # # # # # # # # def prune_columns(df: pd.DataFrame, drop_unnamed=True, null_threshold=0.90, drop_pattern=None):
# # # # # # # # #     drop = []
# # # # # # # # #     for c in df.columns:
# # # # # # # # #         name = str(c).strip()
# # # # # # # # #         if drop_unnamed and (name.lower().startswith("unnamed") or name == "" or name.lower() == "nan"):
# # # # # # # # #             drop.append(c)
# # # # # # # # #             continue
# # # # # # # # #         if drop_pattern and re.match(drop_pattern, name):
# # # # # # # # #             drop.append(c)
# # # # # # # # #             continue
# # # # # # # # #         if df[c].isna().mean() >= null_threshold:
# # # # # # # # #             drop.append(c)
# # # # # # # # #             continue
# # # # # # # # #         if df[c].dtype == "object":
# # # # # # # # #             s = df[c].astype(str).str.strip().replace("nan", "")
# # # # # # # # #             if s.replace("", np.nan).dropna().empty:
# # # # # # # # #                 drop.append(c)
# # # # # # # # #                 continue
# # # # # # # # #     return df.drop(columns=drop), drop

# # # # # # # # # def arrow_safe(df: pd.DataFrame) -> pd.DataFrame:
# # # # # # # # #     out = df.copy()
# # # # # # # # #     for c in out.columns:
# # # # # # # # #         if out[c].dtype == "object":
# # # # # # # # #             out[c] = out[c].astype("string")
# # # # # # # # #     return out

# # # # # # # # # def reduce_ticks(labels, max_ticks=25):
# # # # # # # # #     n = len(labels)
# # # # # # # # #     if n <= max_ticks:
# # # # # # # # #         return list(range(n)), labels
# # # # # # # # #     step = max(1, n // max_ticks)
# # # # # # # # #     idxs = list(range(0, n, step))
# # # # # # # # #     return idxs, [labels[i] for i in idxs]

# # # # # # # # # def hex_to_rgba(hex_color: str, alpha: float) -> str:
# # # # # # # # #     try:
# # # # # # # # #         hc = str(hex_color).lstrip("#")
# # # # # # # # #         if len(hc) != 6: return hex_color
# # # # # # # # #         r = int(hc[0:2], 16); g = int(hc[2:4], 16); b = int(hc[4:6], 16)
# # # # # # # # #         return f"rgba({r},{g},{b},{alpha})"
# # # # # # # # #     except:
# # # # # # # # #         return hex_color

# # # # # # # # # def palette_color(palette: list, idx: int) -> str:
# # # # # # # # #     return palette[idx % len(palette)]

# # # # # # # # # PALETTES = {
# # # # # # # # #     "Plotly": px.colors.qualitative.Plotly,
# # # # # # # # #     "D3": px.colors.qualitative.D3,
# # # # # # # # #     "Bold": px.colors.qualitative.Bold,
# # # # # # # # #     "Dark24": px.colors.qualitative.Dark24,
# # # # # # # # #     "G10": px.colors.qualitative.G10,
# # # # # # # # #     "Alphabet": px.colors.qualitative.Alphabet,
# # # # # # # # # }

# # # # # # # # # # ---------------- Main App ----------------
# # # # # # # # # uploaded_files = st.file_uploader("Upload Excel file(s) (.xlsx)", type=["xlsx"], accept_multiple_files=True)

# # # # # # # # # if uploaded_files:
# # # # # # # # #     all_dfs = []
# # # # # # # # #     for uploaded in uploaded_files:
# # # # # # # # #         try:
# # # # # # # # #             xls, engine_used = load_excel(uploaded)
# # # # # # # # #             sheet = find_sheet(xls)
# # # # # # # # #             raw = pd.read_excel(uploaded, sheet_name=sheet, engine=engine_used, header=None)
# # # # # # # # #             df = header_detect_clean(raw)
# # # # # # # # #             df, _ = prune_columns(df, drop_unnamed=True, null_threshold=0.90, drop_pattern=r"^0\.0_.*")
# # # # # # # # #             df['Source_File'] = uploaded.name
# # # # # # # # #             all_dfs.append(df)
# # # # # # # # #         except Exception as e:
# # # # # # # # #             st.error(f"Error with {uploaded.name}: {e}")

# # # # # # # # #     if not all_dfs:
# # # # # # # # #         st.stop()

# # # # # # # # #     df_combined = pd.concat(all_dfs, ignore_index=True)

# # # # # # # # #     num_cols = df_combined.select_dtypes(include="number").columns.tolist()
# # # # # # # # #     valid_y_cols = [c for c in num_cols 
# # # # # # # # #                     if pd.to_numeric(df_combined[c], errors="coerce").dropna().abs().sum() != 0]

# # # # # # # # #     cat_cols = [c for c in df_combined.columns if c not in num_cols and c != 'Source_File']

# # # # # # # # #     # ---------------- SIDEBAR ----------------
# # # # # # # # #     with st.sidebar:
# # # # # # # # #         st.header("Controls")
# # # # # # # # #         x_col = st.selectbox("X-axis", options=cat_cols or [None])
# # # # # # # # #         y_cols = st.multiselect("Y-axis (numeric)", options=valid_y_cols, 
# # # # # # # # #                                default=valid_y_cols[:1] if valid_y_cols else [])

# # # # # # # # #         agg = st.radio("Aggregate by X", ["None", "sum", "mean"], index=0)
# # # # # # # # #         top_n = st.number_input("Top N", min_value=0, value=0, step=1)

# # # # # # # # #         st.markdown("### Colors")
# # # # # # # # #         palette_name = st.selectbox("Palette", list(PALETTES.keys()), index=0)
# # # # # # # # #         palette = PALETTES[palette_name]
# # # # # # # # #         use_custom_colors = st.checkbox("Use custom colors", value=False)

# # # # # # # # #         custom_color_map = {}
# # # # # # # # #         if use_custom_colors and y_cols:
# # # # # # # # #             for i, y in enumerate(y_cols):
# # # # # # # # #                 color_val = st.color_picker(f"{y}", value=palette_color(palette, i), key=f"col_{y}")
# # # # # # # # #                 custom_color_map[y] = color_val

# # # # # # # # #         st.markdown("**Filters**")
# # # # # # # # #         active_filters = {}
# # # # # # # # #         for c in cat_cols[:6]:
# # # # # # # # #             vals = df_combined[c].dropna().astype(str).unique().tolist()
# # # # # # # # #             if 1 < len(vals) <= 200:
# # # # # # # # #                 sel = st.multiselect(f"Filter {c}", options=vals, default=[], placeholder="(all)")
# # # # # # # # #                 if sel:
# # # # # # # # #                     active_filters[c] = set(sel)

# # # # # # # # #         chart_type = st.radio("Chart type", ["Line", "Bar", "Area"], index=1)
# # # # # # # # #         use_log = st.checkbox("Log scale (Y)", value=False)
# # # # # # # # #         rolling = st.number_input("Rolling mean window", min_value=0, value=0, step=1)
# # # # # # # # #         max_ticks = st.slider("Max X ticks", 5, 60, 25)

# # # # # # # # #     # Apply filters & aggregation (same logic)
# # # # # # # # #     filt = pd.Series(True, index=df_combined.index)
# # # # # # # # #     for col, allowed in active_filters.items():
# # # # # # # # #         filt &= df_combined[col].astype(str).isin(allowed)
# # # # # # # # #     df_f = df_combined.loc[filt].copy()

# # # # # # # # #     if x_col and agg != "None" and y_cols:
# # # # # # # # #         df_f = df_f.groupby(x_col, dropna=False)[y_cols].agg(agg).reset_index()

# # # # # # # # #     if top_n and y_cols and x_col:
# # # # # # # # #         df_f = df_f.sort_values(y_cols[0], ascending=False).head(int(top_n))

# # # # # # # # #     if y_cols:
# # # # # # # # #         y_num = df_f[y_cols].apply(pd.to_numeric, errors="coerce").fillna(0)
# # # # # # # # #         df_f = df_f[y_num.abs().sum(axis=1) != 0].reset_index(drop=True)

# # # # # # # # #     if df_f.empty:
# # # # # # # # #         st.warning("No data left after filtering.")
# # # # # # # # #         st.stop()

# # # # # # # # #     st.subheader("Cleaned & Filtered Data")
# # # # # # # # #     st.dataframe(arrow_safe(df_f), use_container_width=True, height=400)

# # # # # # # # #     if not y_cols:
# # # # # # # # #         st.stop()

# # # # # # # # #     # ==================== MAIN CHARTS (unchanged) ====================
# # # # # # # # #     st.subheader("Main Charts")
# # # # # # # # #     x_vals = df_f[x_col].astype(str) if x_col else pd.Series(range(len(df_f)), name="Index")

# # # # # # # # #     for i, y in enumerate(y_cols):
# # # # # # # # #         color = custom_color_map.get(y) if use_custom_colors else palette_color(palette, i)
# # # # # # # # #         df_plot = pd.DataFrame({x_col or "Index": x_vals, y: pd.to_numeric(df_f[y], errors="coerce").fillna(0)})

# # # # # # # # #         if chart_type == "Bar":
# # # # # # # # #             fig = px.bar(df_plot, x=df_plot.columns[0], y=y)
# # # # # # # # #             fig.update_traces(marker_color=color)
# # # # # # # # #         elif chart_type == "Line":
# # # # # # # # #             fig = px.line(df_plot, x=df_plot.columns[0], y=y)
# # # # # # # # #             fig.update_traces(line_color=color)
# # # # # # # # #         else:
# # # # # # # # #             fig = px.area(df_plot, x=df_plot.columns[0], y=y)
# # # # # # # # #             fig.update_traces(line_color=color, fillcolor=hex_to_rgba(color, 0.25))

# # # # # # # # #         if rolling > 1:
# # # # # # # # #             roll = df_plot[y].rolling(rolling, min_periods=1).mean()
# # # # # # # # #             fig.add_scatter(x=df_plot.iloc[:,0], y=roll, mode="lines",
# # # # # # # # #                             name=f"Rolling", line=dict(color=hex_to_rgba(color, 0.8), dash="dash"))

# # # # # # # # #         idxs, labels = reduce_ticks(df_plot.iloc[:,0].tolist(), max_ticks)
# # # # # # # # #         fig.update_layout(
# # # # # # # # #             xaxis=dict(tickmode="array", tickvals=[df_plot.iloc[i,0] for i in idxs], ticktext=labels),
# # # # # # # # #             yaxis_type="log" if use_log else "linear",
# # # # # # # # #             height=480,
# # # # # # # # #             title=f"{y} - {chart_type}"
# # # # # # # # #         )
# # # # # # # # #         st.plotly_chart(fig, use_container_width=True)

# # # # # # # # #     # ==================== IMPROVED MONTH-WISE ANALYSIS ====================
# # # # # # # # #     st.subheader("📅 Month-wise Analysis")

# # # # # # # # #     # Improved detection: look for date-like column headers (2023-04-01, 2023-05, etc.)
# # # # # # # # #     month_col = None
# # # # # # # # #     date_pattern = re.compile(r'^\d{4}-\d{1,2}(-\d{1,2})?')

# # # # # # # # #     for col in df_f.columns:
# # # # # # # # #         col_str = str(col).strip()
# # # # # # # # #         if date_pattern.match(col_str) or any(x in col_str.lower() for x in ['month', 'period', 'date']):
# # # # # # # # #             month_col = col
# # # # # # # # #             break

# # # # # # # # #     if month_col:
# # # # # # # # #         df_month = df_f.copy()
        
# # # # # # # # #         # Convert column name or values to clean YYYY-MM format
# # # # # # # # #         try:
# # # # # # # # #             # If the column itself is a date string (like header "2023-04-01")
# # # # # # # # #             month_key = pd.to_datetime(str(month_col), errors='coerce')
# # # # # # # # #             if pd.notna(month_key):
# # # # # # # # #                 df_month['Month'] = month_key.strftime('%Y-%m')
# # # # # # # # #             else:
# # # # # # # # #                 # Try parsing values inside the column
# # # # # # # # #                 df_month['Month'] = pd.to_datetime(df_month[month_col], errors='coerce').dt.strftime('%Y-%m')
# # # # # # # # #         except:
# # # # # # # # #             df_month['Month'] = df_month[month_col].astype(str).str[:7]  # take first 7 chars (YYYY-MM)

# # # # # # # # #         df_month = df_month.dropna(subset=['Month'])
        
# # # # # # # # #         if not df_month.empty and y_cols:
# # # # # # # # #             monthly_sum = df_month.groupby('Month')[y_cols].sum().reset_index()
# # # # # # # # #             monthly_sum = monthly_sum.sort_values('Month')

# # # # # # # # #             st.success(f"✅ Using **{month_col}** for month-wise view")

# # # # # # # # #             for y in y_cols:
# # # # # # # # #                 fig_m = px.bar(monthly_sum, x='Month', y=y,
# # # # # # # # #                               title=f"Month-wise {y} (Total)",
# # # # # # # # #                               labels={'Month': 'Month'})
# # # # # # # # #                 fig_m.update_traces(marker_color=palette_color(palette, y_cols.index(y)))
# # # # # # # # #                 fig_m.update_layout(height=450)
# # # # # # # # #                 st.plotly_chart(fig_m, use_container_width=True)
# # # # # # # # #         else:
# # # # # # # # #             st.info("No data available for month-wise after grouping.")
# # # # # # # # #     else:
# # # # # # # # #         st.info("Still could not detect a month/date column. "
# # # # # # # # #                 "Please rename the column containing months (e.g. 2023-04-01) to **'Month'** and re-upload.")

# # # # # # # # #     # Download
# # # # # # # # #     st.download_button(
# # # # # # # # #         "Download cleaned/filtered data (CSV)",
# # # # # # # # #         df_f.to_csv(index=False).encode("utf-8"),
# # # # # # # # #         file_name="ghg_summary_cleaned.csv",
# # # # # # # # #         mime="text/csv"
# # # # # # # # #     )

# # # # # # # # # else:
# # # # # # # # #     st.info("Upload your Excel file(s) to begin.")


# # # # # # # # import streamlit as st
# # # # # # # # import pandas as pd
# # # # # # # # import numpy as np
# # # # # # # # import plotly.express as px
# # # # # # # # import plotly.graph_objects as go
# # # # # # # # from openpyxl import load_workbook
# # # # # # # # import datetime

# # # # # # # # st.set_page_config(page_title="GHG Summary Analyzer", layout="wide")
# # # # # # # # st.title("GHG Summary Analyzer")

# # # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # # # CONSTANTS
# # # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # # BASELINE_COLS     = list(range(7, 19))
# # # # # # # # ASSESSMENT_COLS   = list(range(20, 32))
# # # # # # # # BASELINE_MONTHS   = [f"2023-{m:02d}" for m in range(4,13)] + [f"2024-{m:02d}" for m in range(1,4)]
# # # # # # # # ASSESSMENT_MONTHS = [f"2025-{m:02d}" for m in range(4,13)] + [f"2026-{m:02d}" for m in range(1,4)]

# # # # # # # # GROUP_LABELS = {"A","B 1","B 2","C","C.1","C.2","C.2.1","D","D1","D2"}

# # # # # # # # SECTION_PARENTS = {
# # # # # # # #     "A1":"A","A2":"A","A11":"A","A12":"A",
# # # # # # # #     **{f"B1.{i}":"B1" for i in [1,2,3,4,5,6,7,8,9,10]},
# # # # # # # #     **{f"B2.{i}":"B2" for i in [1,2,3,4,5,6,7,8,9,10]},
# # # # # # # # }

# # # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # # # PARSER
# # # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # # def parse_excel(uploaded_file):
# # # # # # # #     wb = load_workbook(uploaded_file, read_only=True, data_only=True)
# # # # # # # #     ws = wb.active
# # # # # # # #     rows = list(ws.iter_rows(values_only=True))

# # # # # # # #     current_parent = "Other"
# # # # # # # #     current_group_label = "Other"   # track D1, D2, C.1, C.2 etc. for (i),(ii) disambiguation
# # # # # # # #     records = []

# # # # # # # #     for row in rows:
# # # # # # # #         if all(v is None for v in row):
# # # # # # # #             continue
# # # # # # # #         if any(isinstance(v, datetime.datetime) for v in row):
# # # # # # # #             continue

# # # # # # # #         raw_sec  = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
# # # # # # # #         desc     = str(row[3]).strip() if len(row) > 3 and row[3] is not None else ""
# # # # # # # #         freq     = str(row[4]).strip() if len(row) > 4 and row[4] is not None else ""
# # # # # # # #         units    = str(row[5]).strip() if len(row) > 5 and row[5] is not None else ""

# # # # # # # #         # Track parent group
# # # # # # # #         if raw_sec in GROUP_LABELS:
# # # # # # # #             current_parent = raw_sec
# # # # # # # #             current_group_label = raw_sec
# # # # # # # #             continue

# # # # # # # #         if not raw_sec or raw_sec.lower() in ("none","nan"):
# # # # # # # #             continue
# # # # # # # #         if desc.lower() in ("none","nan",""):
# # # # # # # #             continue

# # # # # # # #         parent = SECTION_PARENTS.get(raw_sec, current_parent)
# # # # # # # #         # For (i),(ii) etc. sub-rows, tag with their group for uniqueness
# # # # # # # #         display_sec = f"{current_group_label}>{raw_sec}" if raw_sec.startswith("(") else raw_sec

# # # # # # # #         for col_idx in BASELINE_COLS + ASSESSMENT_COLS:
# # # # # # # #             val = row[col_idx] if col_idx < len(row) else None
# # # # # # # #             if col_idx in BASELINE_COLS:
# # # # # # # #                 month_str  = BASELINE_MONTHS[col_idx - 7]
# # # # # # # #                 year_label = "Baseline 2023-24"
# # # # # # # #             else:
# # # # # # # #                 month_str  = ASSESSMENT_MONTHS[col_idx - 20]
# # # # # # # #                 year_label = "Assessment 2025-26"

# # # # # # # #             try:
# # # # # # # #                 numeric_val = float(val)
# # # # # # # #             except (TypeError, ValueError):
# # # # # # # #                 numeric_val = np.nan

# # # # # # # #             records.append({
# # # # # # # #                 "Parent":      parent,
# # # # # # # #                 "Section":     display_sec,
# # # # # # # #                 "RawSection":  raw_sec,
# # # # # # # #                 "Description": desc,
# # # # # # # #                 "Frequency":   freq,
# # # # # # # #                 "Units":       units,
# # # # # # # #                 "Year":        year_label,
# # # # # # # #                 "Month":       month_str,
# # # # # # # #                 "Value":       numeric_val,
# # # # # # # #             })

# # # # # # # #     df = pd.DataFrame(records)
# # # # # # # #     df["Value"] = df["Value"].fillna(0)
# # # # # # # #     return df


# # # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # # # UPLOAD
# # # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # # uploaded_files = st.file_uploader(
# # # # # # # #     "Upload Excel file(s) (.xlsx)", type=["xlsx"], accept_multiple_files=True
# # # # # # # # )
# # # # # # # # if not uploaded_files:
# # # # # # # #     st.info("Upload your Excel file(s) to begin.")
# # # # # # # #     st.stop()

# # # # # # # # all_dfs = []
# # # # # # # # for f in uploaded_files:
# # # # # # # #     try:
# # # # # # # #         parsed = parse_excel(f)
# # # # # # # #         parsed["Source_File"] = f.name
# # # # # # # #         all_dfs.append(parsed)
# # # # # # # #     except Exception as e:
# # # # # # # #         st.error(f"Error reading {f.name}: {e}")

# # # # # # # # if not all_dfs:
# # # # # # # #     st.stop()

# # # # # # # # df = pd.concat(all_dfs, ignore_index=True)
# # # # # # # # all_months = sorted(df["Month"].unique().tolist())

# # # # # # # # # Frequency/type options for dropdown (col 4 meaningful values)
# # # # # # # # freq_options = sorted([
# # # # # # # #     v for v in df["Frequency"].unique()
# # # # # # # #     if v and v.lower() not in ("none","nan","")
# # # # # # # # ])

# # # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # # # TABS
# # # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # # tab1, tab2 = st.tabs(["📊 Data & Charts", "🔮 Assessment Prediction"])

# # # # # # # # # ══════════════════════════════════════════════════════════════════════════════
# # # # # # # # # TAB 1
# # # # # # # # # ══════════════════════════════════════════════════════════════════════════════
# # # # # # # # with tab1:
# # # # # # # #     with st.sidebar:
# # # # # # # #         st.header("Filters")
# # # # # # # #         sel_freq = st.selectbox(
# # # # # # # #             "Type",
# # # # # # # #             options=["All"] + freq_options,
# # # # # # # #             index=0,
# # # # # # # #             key="t1_freq"
# # # # # # # #         )
# # # # # # # #         sel_months = st.multiselect(
# # # # # # # #             "Months",
# # # # # # # #             options=all_months,
# # # # # # # #             default=all_months,
# # # # # # # #             key="t1_months"
# # # # # # # #         )
# # # # # # # #         chart_type = st.radio("Chart type", ["Bar", "Line"], index=0, key="t1_chart")

# # # # # # # #     df1 = df[df["Month"].isin(sel_months)].copy()
# # # # # # # #     if sel_freq != "All":
# # # # # # # #         df1 = df1[df1["Frequency"] == sel_freq]

# # # # # # # #     if df1.empty:
# # # # # # # #         st.warning("No data for selected filters.")
# # # # # # # #         st.stop()

# # # # # # # #     # Pivot table
# # # # # # # #     st.subheader("Data Table — All Sections")
# # # # # # # #     pivot = (
# # # # # # # #         df1.pivot_table(
# # # # # # # #             index=["Parent","Section","Description","Frequency","Units","Year"],
# # # # # # # #             columns="Month",
# # # # # # # #             values="Value",
# # # # # # # #             aggfunc="sum"
# # # # # # # #         ).reset_index()
# # # # # # # #     )
# # # # # # # #     pivot.columns.name = None
# # # # # # # #     st.dataframe(pivot, use_container_width=True, height=440)

# # # # # # # #     # Charts per parent group
# # # # # # # #     st.subheader("Charts — All Sections")
# # # # # # # #     colors = px.colors.qualitative.D3

# # # # # # # #     for p_idx, parent in enumerate(df1["Parent"].unique().tolist()):
# # # # # # # #         df_p = df1[df1["Parent"] == parent]
# # # # # # # #         with st.expander(f"▶ {parent}", expanded=False):
# # # # # # # #             for s_idx, sec in enumerate(df_p["Section"].unique().tolist()):
# # # # # # # #                 df_sec  = df_p[df_p["Section"] == sec]
# # # # # # # #                 desc    = df_sec["Description"].iloc[0]
# # # # # # # #                 units   = df_sec["Units"].iloc[0]
# # # # # # # #                 title   = f"{sec} — {desc} ({units})"
# # # # # # # #                 df_plot = df_sec.sort_values("Month")
# # # # # # # #                 chart_key = f"t1_{parent}_{sec}_{p_idx}_{s_idx}"

# # # # # # # #                 if chart_type == "Bar":
# # # # # # # #                     fig = px.bar(df_plot, x="Month", y="Value", color="Year",
# # # # # # # #                                  title=title, barmode="group",
# # # # # # # #                                  color_discrete_sequence=colors)
# # # # # # # #                 else:
# # # # # # # #                     fig = px.line(df_plot, x="Month", y="Value", color="Year",
# # # # # # # #                                   title=title, markers=True,
# # # # # # # #                                   color_discrete_sequence=colors)

# # # # # # # #                 fig.update_layout(height=340, margin=dict(t=40,b=30),
# # # # # # # #                                   yaxis_title=units, xaxis_title="Month",
# # # # # # # #                                   legend_title="Year")
# # # # # # # #                 st.plotly_chart(fig, use_container_width=True, key=chart_key)

# # # # # # # # # ══════════════════════════════════════════════════════════════════════════════
# # # # # # # # # TAB 2 — Prediction
# # # # # # # # # ══════════════════════════════════════════════════════════════════════════════
# # # # # # # # with tab2:
# # # # # # # #     st.subheader("Assessment Year Prediction")
# # # # # # # #     st.caption("Uses Baseline monthly data to predict unfilled Assessment months via linear trend.")

# # # # # # # #     baseline_df = df[df["Year"] == "Baseline 2023-24"].copy()
# # # # # # # #     assess_df   = df[df["Year"] == "Assessment 2025-26"].copy()

# # # # # # # #     c1, c2 = st.columns([2, 4])
# # # # # # # #     with c1:
# # # # # # # #         pred_freq = st.selectbox(
# # # # # # # #             "Frequency / Type",
# # # # # # # #             options=["All"] + freq_options,
# # # # # # # #             index=0,
# # # # # # # #             key="pred_freq"
# # # # # # # #         )

# # # # # # # #     base_filtered = baseline_df if pred_freq == "All" else baseline_df[baseline_df["Frequency"] == pred_freq]

# # # # # # # #     valid_indicators = []
# # # # # # # #     for (sec, desc, units), grp in base_filtered.groupby(["Section","Description","Units"]):
# # # # # # # #         if grp[grp["Value"] != 0].shape[0] >= 2:
# # # # # # # #             valid_indicators.append(f"{sec} | {desc} ({units})")

# # # # # # # #     if not valid_indicators:
# # # # # # # #         st.info("No indicators with enough baseline data for the selected type.")
# # # # # # # #         st.stop()

# # # # # # # #     with c2:
# # # # # # # #         chosen = st.selectbox("Select Indicator", options=valid_indicators, key="pred_ind")

# # # # # # # #     sec_chosen   = chosen.split(" | ")[0]
# # # # # # # #     rest         = chosen.split(" | ")[1]
# # # # # # # #     desc_chosen  = rest.rsplit(" (", 1)[0]
# # # # # # # #     units_chosen = rest.rsplit("(", 1)[-1].rstrip(")")

# # # # # # # #     base_data = baseline_df[
# # # # # # # #         (baseline_df["Section"] == sec_chosen) &
# # # # # # # #         (baseline_df["Description"] == desc_chosen)
# # # # # # # #     ].sort_values("Month")[["Month","Value"]].copy()

# # # # # # # #     asmnt_data = assess_df[
# # # # # # # #         (assess_df["Section"] == sec_chosen) &
# # # # # # # #         (assess_df["Description"] == desc_chosen)
# # # # # # # #     ].sort_values("Month")[["Month","Value"]].copy()

# # # # # # # #     filled_months   = asmnt_data[asmnt_data["Value"] != 0]["Month"].tolist()
# # # # # # # #     unfilled_months = asmnt_data[asmnt_data["Value"] == 0]["Month"].tolist()

# # # # # # # #     base_nz = base_data[base_data["Value"] != 0].copy().reset_index(drop=True)
# # # # # # # #     base_nz["idx"] = np.arange(len(base_nz))

# # # # # # # #     if len(base_nz) >= 2:
# # # # # # # #         coeffs = np.polyfit(base_nz["idx"], base_nz["Value"], 1)
# # # # # # # #         poly   = np.poly1d(coeffs)

# # # # # # # #         all_assess = asmnt_data["Month"].tolist()
# # # # # # # #         pred_dict  = {m: max(0, poly(len(base_nz) + i)) for i, m in enumerate(all_assess)}

# # # # # # # #         fig = go.Figure()
# # # # # # # #         fig.add_trace(go.Scatter(
# # # # # # # #             x=base_data["Month"], y=base_data["Value"],
# # # # # # # #             mode="lines+markers", name="Baseline (Actual)",
# # # # # # # #             line=dict(color="#1f77b4", width=2), marker=dict(size=7)
# # # # # # # #         ))
# # # # # # # #         if filled_months:
# # # # # # # #             fd = asmnt_data[asmnt_data["Month"].isin(filled_months)]
# # # # # # # #             fig.add_trace(go.Scatter(
# # # # # # # #                 x=fd["Month"], y=fd["Value"],
# # # # # # # #                 mode="lines+markers", name="Assessment (Actual)",
# # # # # # # #                 line=dict(color="#2ca02c", width=2), marker=dict(size=7)
# # # # # # # #             ))
# # # # # # # #         if unfilled_months:
# # # # # # # #             fig.add_trace(go.Scatter(
# # # # # # # #                 x=unfilled_months,
# # # # # # # #                 y=[pred_dict[m] for m in unfilled_months],
# # # # # # # #                 mode="lines+markers", name="Predicted",
# # # # # # # #                 line=dict(color="#ff7f0e", width=2, dash="dash"),
# # # # # # # #                 marker=dict(size=8, symbol="diamond")
# # # # # # # #             ))

# # # # # # # #         fig.update_layout(
# # # # # # # #             title=f"Prediction: {sec_chosen} — {desc_chosen}",
# # # # # # # #             xaxis_title="Month", yaxis_title=units_chosen,
# # # # # # # #             height=420, margin=dict(t=50, b=40),
# # # # # # # #             legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
# # # # # # # #         )
# # # # # # # #         st.plotly_chart(fig, use_container_width=True, key="pred_main_chart")

# # # # # # # #         if unfilled_months:
# # # # # # # #             st.markdown("**Predicted values for unfilled Assessment months:**")
# # # # # # # #             pred_tbl = pd.DataFrame({
# # # # # # # #                 "Month": unfilled_months,
# # # # # # # #                 f"Predicted ({units_chosen})": [round(pred_dict[m], 3) for m in unfilled_months]
# # # # # # # #             })
# # # # # # # #             st.dataframe(pred_tbl, use_container_width=True, hide_index=True)
# # # # # # # #         else:
# # # # # # # #             st.success("All Assessment months already have data — no prediction needed.")

# # # # # # # #         direction = "↑ Increasing" if coeffs[0] > 0 else "↓ Decreasing"
# # # # # # # #         st.info(f"Trend: **{direction}** | Monthly change ≈ **{coeffs[0]:+.3f} {units_chosen}/month**")
# # # # # # # #     else:
# # # # # # # #         st.warning("Not enough non-zero baseline data to build a trend.")

# # # # # # # # # ── Download ──────────────────────────────────────────────────────────────────
# # # # # # # # st.download_button(
# # # # # # # #     "⬇ Download full data (CSV)",
# # # # # # # #     df.to_csv(index=False).encode("utf-8"),
# # # # # # # #     file_name="ghg_summary.csv",
# # # # # # # #     mime="text/csv"
# # # # # # # # )


# # # # # # # import streamlit as st
# # # # # # # import pandas as pd
# # # # # # # import numpy as np
# # # # # # # import plotly.express as px
# # # # # # # import plotly.graph_objects as go
# # # # # # # from openpyxl import load_workbook
# # # # # # # import datetime

# # # # # # # st.set_page_config(page_title="GHG Summary Analyzer", layout="wide")
# # # # # # # st.title("GHG Summary Analyzer")

# # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # # CONSTANTS
# # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # BASELINE_COLS     = list(range(7, 19))
# # # # # # # ASSESSMENT_COLS   = list(range(20, 32))
# # # # # # # BASELINE_MONTHS   = [f"2023-{m:02d}" for m in range(4,13)] + [f"2024-{m:02d}" for m in range(1,4)]
# # # # # # # ASSESSMENT_MONTHS = [f"2025-{m:02d}" for m in range(4,13)] + [f"2026-{m:02d}" for m in range(1,4)]

# # # # # # # GROUP_LABELS = {"A","B 1","B 2","C","C.1","C.2","C.2.1","D","D1","D2"}

# # # # # # # SECTION_PARENTS = {
# # # # # # #     "A1":"A","A2":"A","A11":"A","A12":"A",
# # # # # # #     **{f"B1.{i}":"B1" for i in [1,2,3,4,5,6,7,8,9,10]},
# # # # # # #     **{f"B2.{i}":"B2" for i in [1,2,3,4,5,6,7,8,9,10]},
# # # # # # # }

# # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # # PARSER
# # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # def parse_excel(uploaded_file):
# # # # # # #     wb = load_workbook(uploaded_file, read_only=True, data_only=True)
# # # # # # #     ws = wb.active
# # # # # # #     rows = list(ws.iter_rows(values_only=True))

# # # # # # #     current_parent = "Other"
# # # # # # #     current_group_label = "Other"   # track D1, D2, C.1, C.2 etc. for (i),(ii) disambiguation
# # # # # # #     records = []

# # # # # # #     for row in rows:
# # # # # # #         if all(v is None for v in row):
# # # # # # #             continue
# # # # # # #         if any(isinstance(v, datetime.datetime) for v in row):
# # # # # # #             continue

# # # # # # #         raw_sec  = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
# # # # # # #         desc     = str(row[3]).strip() if len(row) > 3 and row[3] is not None else ""
# # # # # # #         freq     = str(row[4]).strip() if len(row) > 4 and row[4] is not None else ""
# # # # # # #         units    = str(row[5]).strip() if len(row) > 5 and row[5] is not None else ""

# # # # # # #         # Track parent group
# # # # # # #         if raw_sec in GROUP_LABELS:
# # # # # # #             current_parent = raw_sec
# # # # # # #             current_group_label = raw_sec
# # # # # # #             continue

# # # # # # #         if not raw_sec or raw_sec.lower() in ("none","nan"):
# # # # # # #             continue
# # # # # # #         if desc.lower() in ("none","nan",""):
# # # # # # #             continue

# # # # # # #         parent = SECTION_PARENTS.get(raw_sec, current_parent)
# # # # # # #         # For (i),(ii) etc. sub-rows, tag with their group for uniqueness
# # # # # # #         display_sec = f"{current_group_label}>{raw_sec}" if raw_sec.startswith("(") else raw_sec

# # # # # # #         for col_idx in BASELINE_COLS + ASSESSMENT_COLS:
# # # # # # #             val = row[col_idx] if col_idx < len(row) else None
# # # # # # #             if col_idx in BASELINE_COLS:
# # # # # # #                 month_str  = BASELINE_MONTHS[col_idx - 7]
# # # # # # #                 year_label = "Baseline 2023-24"
# # # # # # #             else:
# # # # # # #                 month_str  = ASSESSMENT_MONTHS[col_idx - 20]
# # # # # # #                 year_label = "Assessment 2025-26"

# # # # # # #             try:
# # # # # # #                 numeric_val = float(val)
# # # # # # #             except (TypeError, ValueError):
# # # # # # #                 numeric_val = np.nan

# # # # # # #             records.append({
# # # # # # #                 "Parent":      parent,
# # # # # # #                 "Section":     display_sec,
# # # # # # #                 "RawSection":  raw_sec,
# # # # # # #                 "Description": desc,
# # # # # # #                 "Frequency":   freq,
# # # # # # #                 "Units":       units,
# # # # # # #                 "Year":        year_label,
# # # # # # #                 "Month":       month_str,
# # # # # # #                 "Value":       numeric_val,
# # # # # # #             })

# # # # # # #     df = pd.DataFrame(records)
# # # # # # #     df["Value"] = df["Value"].fillna(0)
# # # # # # #     return df


# # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # # UPLOAD
# # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # uploaded_files = st.file_uploader(
# # # # # # #     "Upload Excel file(s) (.xlsx)", type=["xlsx"], accept_multiple_files=True
# # # # # # # )
# # # # # # # if not uploaded_files:
# # # # # # #     st.info("Upload your Excel file(s) to begin.")
# # # # # # #     st.stop()

# # # # # # # all_dfs = []
# # # # # # # for f in uploaded_files:
# # # # # # #     try:
# # # # # # #         parsed = parse_excel(f)
# # # # # # #         parsed["Source_File"] = f.name
# # # # # # #         all_dfs.append(parsed)
# # # # # # #     except Exception as e:
# # # # # # #         st.error(f"Error reading {f.name}: {e}")

# # # # # # # if not all_dfs:
# # # # # # #     st.stop()

# # # # # # # df = pd.concat(all_dfs, ignore_index=True)
# # # # # # # all_months = sorted(df["Month"].unique().tolist())

# # # # # # # # Frequency/type options for dropdown (col 4 meaningful values)
# # # # # # # freq_options = sorted([
# # # # # # #     v for v in df["Frequency"].unique()
# # # # # # #     if v and v.lower() not in ("none","nan","")
# # # # # # # ])

# # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # # TABS
# # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # tab1, tab2 = st.tabs(["📊 Data & Charts", "🔮 Assessment Prediction"])

# # # # # # # freq_display_options = []

# # # # # # # for (sec, desc, units), grp in df.groupby(["Section", "Description", "Units"]):
# # # # # # #     freq_display_options.append((sec, desc, units))
# # # # # # # # ══════════════════════════════════════════════════════════════════════════════
# # # # # # # # TAB 1
# # # # # # # # ══════════════════════════════════════════════════════════════════════════════
# # # # # # # with tab1:
# # # # # # #     with st.sidebar:
# # # # # # #         st.header("Filters")
# # # # # # #         sel_freq = st.selectbox(
# # # # # # #             "Type",
# # # # # # #             options=["All"] + freq_options,
# # # # # # #             index=0,
# # # # # # #             key="t1_freq"
# # # # # # #         )
# # # # # # #         sel_months = st.multiselect(
# # # # # # #             "Months",
# # # # # # #             options=all_months,
# # # # # # #             default=all_months,
# # # # # # #             key="t1_months"
# # # # # # #         )
# # # # # # #         chart_type = st.radio("Chart type", ["Bar", "Line"], index=0, key="t1_chart")

# # # # # # #     df1 = df[df["Month"].isin(sel_months)].copy()
# # # # # # #     if sel_freq != "All":
# # # # # # #         df1 = df1[df1["Frequency"] == sel_freq]

# # # # # # #     if df1.empty:
# # # # # # #         st.warning("No data for selected filters.")
# # # # # # #         st.stop()

# # # # # # #     # Pivot table
# # # # # # #     st.subheader("Data Table — All Sections")
# # # # # # #     pivot = (
# # # # # # #         df1.pivot_table(
# # # # # # #             index=["Parent","Section","Description","Frequency","Units","Year"],
# # # # # # #             columns="Month",
# # # # # # #             values="Value",
# # # # # # #             aggfunc="sum"
# # # # # # #         ).reset_index()
# # # # # # #     )
# # # # # # #     pivot.columns.name = None
# # # # # # #     st.dataframe(pivot, use_container_width=True, height=440)

# # # # # # #     # Charts per parent group
# # # # # # #     st.subheader("Charts — All Sections")
# # # # # # #     colors = px.colors.qualitative.D3

# # # # # # #     for p_idx, parent in enumerate(df1["Parent"].unique().tolist()):
# # # # # # #         df_p = df1[df1["Parent"] == parent]
# # # # # # #         with st.expander(f"▶ {parent}", expanded=False):
# # # # # # #             for s_idx, sec in enumerate(df_p["Section"].unique().tolist()):
# # # # # # #                 df_sec  = df_p[df_p["Section"] == sec]
# # # # # # #                 desc    = df_sec["Description"].iloc[0]
# # # # # # #                 units   = df_sec["Units"].iloc[0]
# # # # # # #                 title   = f"{sec} — {desc} ({units})"
# # # # # # #                 df_plot = df_sec.sort_values("Month")
# # # # # # #                 chart_key = f"t1_{parent}_{sec}_{p_idx}_{s_idx}"

# # # # # # #                 if chart_type == "Bar":
# # # # # # #                     fig = px.bar(df_plot, x="Month", y="Value", color="Year",
# # # # # # #                                  title=title, barmode="group",
# # # # # # #                                  color_discrete_sequence=colors)
# # # # # # #                 else:
# # # # # # #                     fig = px.line(df_plot, x="Month", y="Value", color="Year",
# # # # # # #                                   title=title, markers=True,
# # # # # # #                                   color_discrete_sequence=colors)

# # # # # # #                 fig.update_layout(height=340, margin=dict(t=40,b=30),
# # # # # # #                                   yaxis_title=units, xaxis_title="Month",
# # # # # # #                                   legend_title="Year")
# # # # # # #                 st.plotly_chart(fig, use_container_width=True, key=chart_key)

# # # # # # # # ══════════════════════════════════════════════════════════════════════════════
# # # # # # # # TAB 2 — Prediction
# # # # # # # # ══════════════════════════════════════════════════════════════════════════════
# # # # # # # with tab2:
# # # # # # #     st.subheader("Assessment Year Prediction")
# # # # # # #     st.caption("Uses Baseline monthly data to predict unfilled Assessment months via linear trend.")

# # # # # # #     baseline_df = df[df["Year"] == "Baseline 2023-24"].copy()
# # # # # # #     assess_df   = df[df["Year"] == "Assessment 2025-26"].copy()

# # # # # # #     # c1, c2 = st.columns([2, 4])
# # # # # # #     # with c1:
# # # # # # #     #     pred_freq = st.selectbox(
# # # # # # #     #         "Frequency / Type",
# # # # # # #     #         options=["All"] + freq_options,
# # # # # # #     #         index=0,
# # # # # # #     #         key="pred_freq"
# # # # # # #     #     )

# # # # # # #     # base_filtered = baseline_df if pred_freq == "All" else baseline_df[baseline_df["Frequency"] == pred_freq]
# # # # # # #     base_filtered = baseline_df.copy()

# # # # # # #     valid_indicators = []
# # # # # # #     for (sec, desc, units), grp in base_filtered.groupby(["Section","Description","Units"]):
# # # # # # #         if grp[grp["Value"] != 0].shape[0] >= 2:
# # # # # # #             valid_indicators.append(f"{sec} | {desc} ({units})")

# # # # # # #     if not valid_indicators:
# # # # # # #         st.info("No indicators with enough baseline data for the selected type.")
# # # # # # #         st.stop()

# # # # # # #     # with c2:
# # # # # # #     #     chosen = st.selectbox("Select Indicator", options=valid_indicators, key="pred_ind")
# # # # # # #     chosen = st.selectbox("Select Indicator", options=valid_indicators, key="pred_ind")

# # # # # # #     sec_chosen   = chosen.split(" | ")[0]
# # # # # # #     rest         = chosen.split(" | ")[1]
# # # # # # #     desc_chosen  = rest.rsplit(" (", 1)[0]
# # # # # # #     units_chosen = rest.rsplit("(", 1)[-1].rstrip(")")

# # # # # # #     base_data = baseline_df[
# # # # # # #         (baseline_df["Section"] == sec_chosen) &
# # # # # # #         (baseline_df["Description"] == desc_chosen)
# # # # # # #     ].sort_values("Month")[["Month","Value"]].copy()

# # # # # # #     asmnt_data = assess_df[
# # # # # # #         (assess_df["Section"] == sec_chosen) &
# # # # # # #         (assess_df["Description"] == desc_chosen)
# # # # # # #     ].sort_values("Month")[["Month","Value"]].copy()

# # # # # # #     filled_months   = asmnt_data[asmnt_data["Value"] != 0]["Month"].tolist()
# # # # # # #     unfilled_months = asmnt_data[asmnt_data["Value"] == 0]["Month"].tolist()

# # # # # # #     base_nz = base_data[base_data["Value"] != 0].copy().reset_index(drop=True)
# # # # # # #     base_nz["idx"] = np.arange(len(base_nz))

# # # # # # #     if len(base_nz) >= 2:
# # # # # # #         coeffs = np.polyfit(base_nz["idx"], base_nz["Value"], 1)
# # # # # # #         poly   = np.poly1d(coeffs)

# # # # # # #         all_assess = asmnt_data["Month"].tolist()
# # # # # # #         pred_dict  = {m: max(0, poly(len(base_nz) + i)) for i, m in enumerate(all_assess)}

# # # # # # #         # Build unified bar chart: all months, colour-coded by type
# # # # # # #         rows_chart = []
# # # # # # #         for _, r in base_data.iterrows():
# # # # # # #             rows_chart.append({"Month": r["Month"], "Value": r["Value"], "Type": "Baseline"})
# # # # # # #         for _, r in asmnt_data.iterrows():
# # # # # # #             if r["Month"] in filled_months:
# # # # # # #                 rows_chart.append({"Month": r["Month"], "Value": r["Value"], "Type": "Assessment (Actual)"})
# # # # # # #             else:
# # # # # # #                 rows_chart.append({"Month": r["Month"], "Value": pred_dict[r["Month"]], "Type": "Predicted"})

# # # # # # #         df_chart = pd.DataFrame(rows_chart)
# # # # # # #         color_map = {
# # # # # # #             "Baseline":            "#1f77b4",
# # # # # # #             "Assessment (Actual)": "#2ca02c",
# # # # # # #             "Predicted":           "#ff7f0e",
# # # # # # #         }

# # # # # # #         fig = px.bar(
# # # # # # #             df_chart, x="Month", y="Value", color="Type",
# # # # # # #             color_discrete_map=color_map,
# # # # # # #             title=f"{sec_chosen} — {desc_chosen}",
# # # # # # #             barmode="group",
# # # # # # #         )
# # # # # # #         for trace in fig.data:
# # # # # # #             if trace.name == "Predicted":
# # # # # # #                 trace.marker.opacity = 0.75

# # # # # # #         fig.update_layout(
# # # # # # #             height=420, margin=dict(t=50, b=40),
# # # # # # #             yaxis_title=units_chosen, xaxis_title="Month",
# # # # # # #             legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
# # # # # # #         )
# # # # # # #         st.plotly_chart(fig, use_container_width=True, key="pred_main_chart")

# # # # # # #         if unfilled_months:
# # # # # # #             st.markdown("**Predicted values for unfilled Assessment months:**")
# # # # # # #             pred_tbl = pd.DataFrame({
# # # # # # #                 "Month": unfilled_months,
# # # # # # #                 f"Predicted ({units_chosen})": [round(pred_dict[m], 3) for m in unfilled_months]
# # # # # # #             })
# # # # # # #             st.dataframe(pred_tbl, use_container_width=True, hide_index=True)
# # # # # # #         else:
# # # # # # #             st.success("All Assessment months already have data — no prediction needed.")

# # # # # # #         direction    = "Up (Increasing)" if coeffs[0] > 0 else "Down (Decreasing)"
# # # # # # #         slope        = round(coeffs[0], 4)
# # # # # # #         intercept    = round(coeffs[1], 4)
# # # # # # #         avg_baseline = round(float(base_nz["Value"].mean()), 3)
# # # # # # #         n_points     = len(base_nz)

# # # # # # #         with st.expander("How was this prediction calculated?", expanded=True):
# # # # # # #             st.markdown(f"""
# # # # # # # **Method: Linear Regression (Least Squares)**

# # # # # # # The prediction uses the {n_points} non-zero months from the **Baseline year (2023-24)**
# # # # # # # to fit a straight-line trend, then extends that line into the **Assessment year (2025-26)**.

# # # # # # # ---

# # # # # # # **Step-by-step:**

# # # # # # # 1. **Collect baseline data** — took all {n_points} months where the value was not zero.
# # # # # # #    Average baseline value: **{avg_baseline} {units_chosen}**

# # # # # # # 2. **Assign a number to each month** — each month gets an index (0, 1, 2 ... {n_points-1})
# # # # # # #    so the formula can work on plain numbers instead of dates.

# # # # # # # 3. **Fit a straight line** using the formula:
# # # # # # #    `Value = slope x month_index + intercept`
# # # # # # #    Slope = **{slope}** | Intercept = **{intercept}**

# # # # # # # 4. **Extend the line forward** — Assessment months continue the index
# # # # # # #    from {n_points} up to {n_points+11}, giving the predicted values.

# # # # # # # 5. **Floor at zero** — any predicted value below 0 is clamped to 0.

# # # # # # # ---

# # # # # # # **Trend direction: {direction}**
# # # # # # # Each month the value changes by approximately **{slope:+.3f} {units_chosen}**.

# # # # # # # *Note: This is a simple linear extrapolation based on the available baseline months.
# # # # # # # It works well for stable metrics. It does not account for seasonality or sudden
# # # # # # # operational changes.*
# # # # # # # """)
# # # # # # #     else:
# # # # # # #         st.warning("Not enough non-zero baseline data to build a trend.")

# # # # # # # # ── Download ──────────────────────────────────────────────────────────────────
# # # # # # # st.download_button(
# # # # # # #     "⬇ Download full data (CSV)",
# # # # # # #     df.to_csv(index=False).encode("utf-8"),
# # # # # # #     file_name="ghg_summary.csv",
# # # # # # #     mime="text/csv"
# # # # # # # )


# # # # # # # import streamlit as st
# # # # # # # import pandas as pd
# # # # # # # import numpy as np
# # # # # # # import plotly.express as px
# # # # # # # import plotly.graph_objects as go
# # # # # # # from openpyxl import load_workbook
# # # # # # # import datetime

# # # # # # # st.set_page_config(page_title="GHG Summary Analyzer", layout="wide")
# # # # # # # st.title("GHG Summary Analyzer")

# # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # # CONSTANTS
# # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # BASELINE_COLS     = list(range(7, 19))
# # # # # # # ASSESSMENT_COLS   = list(range(20, 32))
# # # # # # # BASELINE_MONTHS   = [f"2023-{m:02d}" for m in range(4,13)] + [f"2024-{m:02d}" for m in range(1,4)]
# # # # # # # ASSESSMENT_MONTHS = [f"2025-{m:02d}" for m in range(4,13)] + [f"2026-{m:02d}" for m in range(1,4)]

# # # # # # # GROUP_LABELS = {"A","B 1","B 2","C","C.1","C.2","C.2.1","D","D1","D2"}

# # # # # # # SECTION_PARENTS = {
# # # # # # #     "A1":"A","A2":"A","A11":"A","A12":"A",
# # # # # # #     **{f"B1.{i}":"B1" for i in [1,2,3,4,5,6,7,8,9,10]},
# # # # # # #     **{f"B2.{i}":"B2" for i in [1,2,3,4,5,6,7,8,9,10]},
# # # # # # # }

# # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # # PARSER
# # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # def parse_excel(uploaded_file):
# # # # # # #     wb = load_workbook(uploaded_file, read_only=True, data_only=True)
# # # # # # #     ws = wb.active
# # # # # # #     rows = list(ws.iter_rows(values_only=True))

# # # # # # #     current_parent = "Other"
# # # # # # #     current_group_label = "Other"
# # # # # # #     records = []

# # # # # # #     for row in rows:
# # # # # # #         if all(v is None for v in row):
# # # # # # #             continue
# # # # # # #         if any(isinstance(v, datetime.datetime) for v in row):
# # # # # # #             continue

# # # # # # #         raw_sec  = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
# # # # # # #         desc     = str(row[3]).strip() if len(row) > 3 and row[3] is not None else ""
# # # # # # #         freq     = str(row[4]).strip() if len(row) > 4 and row[4] is not None else ""
# # # # # # #         units    = str(row[5]).strip() if len(row) > 5 and row[5] is not None else ""

# # # # # # #         if raw_sec in GROUP_LABELS:
# # # # # # #             current_parent = raw_sec
# # # # # # #             current_group_label = raw_sec
# # # # # # #             continue

# # # # # # #         if not raw_sec or raw_sec.lower() in ("none","nan"):
# # # # # # #             continue
# # # # # # #         if desc.lower() in ("none","nan",""):
# # # # # # #             continue

# # # # # # #         parent = SECTION_PARENTS.get(raw_sec, current_parent)
# # # # # # #         display_sec = f"{current_group_label}>{raw_sec}" if raw_sec.startswith("(") else raw_sec

# # # # # # #         for col_idx in BASELINE_COLS + ASSESSMENT_COLS:
# # # # # # #             val = row[col_idx] if col_idx < len(row) else None
# # # # # # #             if col_idx in BASELINE_COLS:
# # # # # # #                 month_str  = BASELINE_MONTHS[col_idx - 7]
# # # # # # #                 year_label = "Baseline 2023-24"
# # # # # # #             else:
# # # # # # #                 month_str  = ASSESSMENT_MONTHS[col_idx - 20]
# # # # # # #                 year_label = "Assessment 2025-26"

# # # # # # #             try:
# # # # # # #                 numeric_val = float(val)
# # # # # # #             except (TypeError, ValueError):
# # # # # # #                 numeric_val = np.nan

# # # # # # #             records.append({
# # # # # # #                 "Parent":      parent,
# # # # # # #                 "Section":     display_sec,
# # # # # # #                 "RawSection":  raw_sec,
# # # # # # #                 "Description": desc,
# # # # # # #                 "Frequency":   freq,
# # # # # # #                 "Units":       units,
# # # # # # #                 "Year":        year_label,
# # # # # # #                 "Month":       month_str,
# # # # # # #                 "Value":       numeric_val,
# # # # # # #             })

# # # # # # #     df = pd.DataFrame(records)
# # # # # # #     df["Value"] = df["Value"].fillna(0)
# # # # # # #     return df


# # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # # UPLOAD
# # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # uploaded_files = st.file_uploader(
# # # # # # #     "Upload Excel file(s) (.xlsx)", type=["xlsx"], accept_multiple_files=True
# # # # # # # )
# # # # # # # if not uploaded_files:
# # # # # # #     st.info("Upload your Excel file(s) to begin.")
# # # # # # #     st.stop()

# # # # # # # all_dfs = []
# # # # # # # for f in uploaded_files:
# # # # # # #     try:
# # # # # # #         parsed = parse_excel(f)
# # # # # # #         parsed["Source_File"] = f.name
# # # # # # #         all_dfs.append(parsed)
# # # # # # #     except Exception as e:
# # # # # # #         st.error(f"Error reading {f.name}: {e}")

# # # # # # # if not all_dfs:
# # # # # # #     st.stop()

# # # # # # # df = pd.concat(all_dfs, ignore_index=True)
# # # # # # # all_months = sorted(df["Month"].unique().tolist())

# # # # # # # # Frequency/type options
# # # # # # # freq_options = sorted([
# # # # # # #     v for v in df["Frequency"].unique()
# # # # # # #     if v and v.lower() not in ("none","nan","")
# # # # # # # ])

# # # # # # # # Units options
# # # # # # # units_options = sorted([
# # # # # # #     v for v in df["Units"].unique()
# # # # # # #     if v and v.lower() not in ("none","nan","")
# # # # # # # ])

# # # # # # # # Build section → description mapping for sidebar display
# # # # # # # sec_desc_map = (
# # # # # # #     df[["Section","Description"]]
# # # # # # #     .drop_duplicates("Section")
# # # # # # #     .set_index("Section")["Description"]
# # # # # # #     .to_dict()
# # # # # # # )

# # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # # TABS
# # # # # # # # ─────────────────────────────────────────────────────────────────────────────
# # # # # # # tab1, tab2 = st.tabs(["📊 Data & Charts", "🔮 Assessment Prediction"])

# # # # # # # freq_display_options = []

# # # # # # # for (sec, desc, units), grp in df.groupby(["Section", "Description", "Units"]):
# # # # # # #     freq_display_options.append((sec, desc, units))

# # # # # # # # ══════════════════════════════════════════════════════════════════════════════
# # # # # # # # TAB 1
# # # # # # # # ══════════════════════════════════════════════════════════════════════════════
# # # # # # # with tab1:
# # # # # # #     with st.sidebar:
# # # # # # #         st.header("Filters")

# # # # # # #         # ── Type filter ──────────────────────────────────────────────────────
# # # # # # #         # sel_freq = st.selectbox(
# # # # # # #         #     "Type",
# # # # # # #         #     options=["All"] + freq_options,
# # # # # # #         #     index=0,
# # # # # # #         #     key="t1_freq"
# # # # # # #         # )

# # # # # # #         # ── Units filter (NEW) ───────────────────────────────────────────────
# # # # # # #         sel_units = st.selectbox(
# # # # # # #             "Units",
# # # # # # #             options=["All"] + units_options,
# # # # # # #             index=0,
# # # # # # #             key="t1_units"
# # # # # # #         )

# # # # # # #         # ── Months filter ────────────────────────────────────────────────────
# # # # # # #         sel_months = st.multiselect(
# # # # # # #             "Months",
# # # # # # #             options=all_months,
# # # # # # #             default=all_months,
# # # # # # #             key="t1_months"
# # # # # # #         )

# # # # # # #         # ── Chart type ───────────────────────────────────────────────────────
# # # # # # #         chart_type = st.radio("Chart type", ["Bar", "Line"], index=0, key="t1_chart")

# # # # # # #         # ── Column C — Section Names ─────────────────────────────────────────
# # # # # # #         st.markdown("---")
# # # # # # #         st.subheader("📋 Section Names (Column C)")

# # # # # # #         # Get unique sections with their descriptions, sorted
# # # # # # #         section_info = (
# # # # # # #             df[["Section","Description","Units"]]
# # # # # # #             .drop_duplicates(subset=["Section"])
# # # # # # #             .sort_values("Section")
# # # # # # #         )

# # # # # # #         # Apply current type / units filter so the list stays in sync
# # # # # # #         filtered_section_info = section_info.copy()
# # # # # # #         # if sel_freq != "All":
# # # # # # #         #     valid_secs = df[df["Frequency"] == sel_freq]["Section"].unique()
# # # # # # #         #     filtered_section_info = filtered_section_info[
# # # # # # #         #         filtered_section_info["Section"].isin(valid_secs)
# # # # # # #         #     ]
# # # # # # #         if sel_units != "All":
# # # # # # #             valid_secs2 = df[df["Units"] == sel_units]["Section"].unique()
# # # # # # #             filtered_section_info = filtered_section_info[
# # # # # # #                 filtered_section_info["Section"].isin(valid_secs2)
# # # # # # #             ]

# # # # # # #         for _, row_s in filtered_section_info.iterrows():
# # # # # # #             sec_label = row_s["Section"]
# # # # # # #             desc_label = row_s["Description"]
# # # # # # #             unit_label = row_s["Units"] if row_s["Units"] and row_s["Units"].lower() not in ("none","nan","") else "—"
# # # # # # #             st.markdown(
# # # # # # #                 f"<div style='margin-bottom:4px;'>"
# # # # # # #                 f"<span style='font-weight:600;color:#1565C0;'>{sec_label}</span>"
# # # # # # #                 f"<br><span style='font-size:0.78em;color:#444;'>{desc_label}</span>"
# # # # # # #                 f"<br><span style='font-size:0.72em;color:#888;'>Unit: {unit_label}</span>"
# # # # # # #                 f"</div>",
# # # # # # #                 unsafe_allow_html=True
# # # # # # #             )

# # # # # # #     # ── Apply filters ─────────────────────────────────────────────────────────
# # # # # # #     df1 = df[df["Month"].isin(sel_months)].copy()
# # # # # # #     # if sel_freq != "All":
# # # # # # #     #     df1 = df1[df1["Frequency"] == sel_freq]
# # # # # # #     if sel_units != "All":
# # # # # # #         df1 = df1[df1["Units"] == sel_units]

# # # # # # #     if df1.empty:
# # # # # # #         st.warning("No data for selected filters.")
# # # # # # #         st.stop()

# # # # # # #     # Pivot table
# # # # # # #     st.subheader("Data Table — All Sections")
# # # # # # #     pivot = (
# # # # # # #         df1.pivot_table(
# # # # # # #             index=["Parent","Section","Description","Frequency","Units","Year"],
# # # # # # #             columns="Month",
# # # # # # #             values="Value",
# # # # # # #             aggfunc="sum"
# # # # # # #         ).reset_index()
# # # # # # #     )
# # # # # # #     pivot.columns.name = None
# # # # # # #     st.dataframe(pivot, use_container_width=True, height=440)

# # # # # # #     # Charts per parent group
# # # # # # #     st.subheader("Charts — All Sections")
# # # # # # #     colors = px.colors.qualitative.D3

# # # # # # #     for p_idx, parent in enumerate(df1["Parent"].unique().tolist()):
# # # # # # #         df_p = df1[df1["Parent"] == parent]
# # # # # # #         with st.expander(f"▶ {parent}", expanded=False):
# # # # # # #             for s_idx, sec in enumerate(df_p["Section"].unique().tolist()):
# # # # # # #                 df_sec  = df_p[df_p["Section"] == sec]
# # # # # # #                 desc    = df_sec["Description"].iloc[0]
# # # # # # #                 units   = df_sec["Units"].iloc[0]
# # # # # # #                 title   = f"{sec} — {desc} ({units})"
# # # # # # #                 df_plot = df_sec.sort_values("Month")
# # # # # # #                 chart_key = f"t1_{parent}_{sec}_{p_idx}_{s_idx}"

# # # # # # #                 if chart_type == "Bar":
# # # # # # #                     fig = px.bar(df_plot, x="Month", y="Value", color="Year",
# # # # # # #                                  title=title, barmode="group",
# # # # # # #                                  color_discrete_sequence=colors)
# # # # # # #                 else:
# # # # # # #                     fig = px.line(df_plot, x="Month", y="Value", color="Year",
# # # # # # #                                   title=title, markers=True,
# # # # # # #                                   color_discrete_sequence=colors)

# # # # # # #                 fig.update_layout(height=340, margin=dict(t=40,b=30),
# # # # # # #                                   yaxis_title=units, xaxis_title="Month",
# # # # # # #                                   legend_title="Year")
# # # # # # #                 st.plotly_chart(fig, use_container_width=True, key=chart_key)

# # # # # # # # ══════════════════════════════════════════════════════════════════════════════
# # # # # # # # TAB 2 — Prediction
# # # # # # # # ══════════════════════════════════════════════════════════════════════════════
# # # # # # # with tab2:
# # # # # # #     st.subheader("Assessment Year Prediction")
# # # # # # #     st.caption("Uses Baseline monthly data to predict unfilled Assessment months via linear trend.")

# # # # # # #     baseline_df = df[df["Year"] == "Baseline 2023-24"].copy()
# # # # # # #     assess_df   = df[df["Year"] == "Assessment 2025-26"].copy()

# # # # # # #     base_filtered = baseline_df.copy()

# # # # # # #     valid_indicators = []
# # # # # # #     for (sec, desc, units), grp in base_filtered.groupby(["Section","Description","Units"]):
# # # # # # #         if grp[grp["Value"] != 0].shape[0] >= 2:
# # # # # # #             valid_indicators.append(f"{sec} | {desc} ({units})")

# # # # # # #     if not valid_indicators:
# # # # # # #         st.info("No indicators with enough baseline data for the selected type.")
# # # # # # #         st.stop()

# # # # # # #     chosen = st.selectbox("Select Indicator", options=valid_indicators, key="pred_ind")

# # # # # # #     sec_chosen   = chosen.split(" | ")[0]
# # # # # # #     rest         = chosen.split(" | ")[1]
# # # # # # #     desc_chosen  = rest.rsplit(" (", 1)[0]
# # # # # # #     units_chosen = rest.rsplit("(", 1)[-1].rstrip(")")

# # # # # # #     base_data = baseline_df[
# # # # # # #         (baseline_df["Section"] == sec_chosen) &
# # # # # # #         (baseline_df["Description"] == desc_chosen)
# # # # # # #     ].sort_values("Month")[["Month","Value"]].copy()

# # # # # # #     asmnt_data = assess_df[
# # # # # # #         (assess_df["Section"] == sec_chosen) &
# # # # # # #         (assess_df["Description"] == desc_chosen)
# # # # # # #     ].sort_values("Month")[["Month","Value"]].copy()

# # # # # # #     filled_months   = asmnt_data[asmnt_data["Value"] != 0]["Month"].tolist()
# # # # # # #     unfilled_months = asmnt_data[asmnt_data["Value"] == 0]["Month"].tolist()

# # # # # # #     base_nz = base_data[base_data["Value"] != 0].copy().reset_index(drop=True)
# # # # # # #     base_nz["idx"] = np.arange(len(base_nz))

# # # # # # #     if len(base_nz) >= 2:
# # # # # # #         coeffs = np.polyfit(base_nz["idx"], base_nz["Value"], 1)
# # # # # # #         poly   = np.poly1d(coeffs)

# # # # # # #         all_assess = asmnt_data["Month"].tolist()
# # # # # # #         pred_dict  = {m: max(0, poly(len(base_nz) + i)) for i, m in enumerate(all_assess)}

# # # # # # #         rows_chart = []
# # # # # # #         for _, r in base_data.iterrows():
# # # # # # #             rows_chart.append({"Month": r["Month"], "Value": r["Value"], "Type": "Baseline"})
# # # # # # #         for _, r in asmnt_data.iterrows():
# # # # # # #             if r["Month"] in filled_months:
# # # # # # #                 rows_chart.append({"Month": r["Month"], "Value": r["Value"], "Type": "Assessment (Actual)"})
# # # # # # #             else:
# # # # # # #                 rows_chart.append({"Month": r["Month"], "Value": pred_dict[r["Month"]], "Type": "Predicted"})

# # # # # # #         df_chart = pd.DataFrame(rows_chart)
# # # # # # #         color_map = {
# # # # # # #             "Baseline":            "#1f77b4",
# # # # # # #             "Assessment (Actual)": "#2ca02c",
# # # # # # #             "Predicted":           "#ff7f0e",
# # # # # # #         }

# # # # # # #         fig = px.bar(
# # # # # # #             df_chart, x="Month", y="Value", color="Type",
# # # # # # #             color_discrete_map=color_map,
# # # # # # #             title=f"{sec_chosen} — {desc_chosen}",
# # # # # # #             barmode="group",
# # # # # # #         )
# # # # # # #         for trace in fig.data:
# # # # # # #             if trace.name == "Predicted":
# # # # # # #                 trace.marker.opacity = 0.75

# # # # # # #         fig.update_layout(
# # # # # # #             height=420, margin=dict(t=50, b=40),
# # # # # # #             yaxis_title=units_chosen, xaxis_title="Month",
# # # # # # #             legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
# # # # # # #         )
# # # # # # #         st.plotly_chart(fig, use_container_width=True, key="pred_main_chart")

# # # # # # #         if unfilled_months:
# # # # # # #             st.markdown("**Predicted values for unfilled Assessment months:**")
# # # # # # #             pred_tbl = pd.DataFrame({
# # # # # # #                 "Month": unfilled_months,
# # # # # # #                 f"Predicted ({units_chosen})": [round(pred_dict[m], 3) for m in unfilled_months]
# # # # # # #             })
# # # # # # #             st.dataframe(pred_tbl, use_container_width=True, hide_index=True)
# # # # # # #         else:
# # # # # # #             st.success("All Assessment months already have data — no prediction needed.")

# # # # # # #         direction    = "Up (Increasing)" if coeffs[0] > 0 else "Down (Decreasing)"
# # # # # # #         slope        = round(coeffs[0], 4)
# # # # # # #         intercept    = round(coeffs[1], 4)
# # # # # # #         avg_baseline = round(float(base_nz["Value"].mean()), 3)
# # # # # # #         n_points     = len(base_nz)

# # # # # # #         with st.expander("How was this prediction calculated?", expanded=True):
# # # # # # #             st.markdown(f"""
# # # # # # # **Method: Linear Regression (Least Squares)**

# # # # # # # The prediction uses the {n_points} non-zero months from the **Baseline year (2023-24)**
# # # # # # # to fit a straight-line trend, then extends that line into the **Assessment year (2025-26)**.

# # # # # # # ---

# # # # # # # **Step-by-step:**

# # # # # # # 1. **Collect baseline data** — took all {n_points} months where the value was not zero.
# # # # # # #    Average baseline value: **{avg_baseline} {units_chosen}**

# # # # # # # 2. **Assign a number to each month** — each month gets an index (0, 1, 2 ... {n_points-1})
# # # # # # #    so the formula can work on plain numbers instead of dates.

# # # # # # # 3. **Fit a straight line** using the formula:
# # # # # # #    `Value = slope x month_index + intercept`
# # # # # # #    Slope = **{slope}** | Intercept = **{intercept}**

# # # # # # # 4. **Extend the line forward** — Assessment months continue the index
# # # # # # #    from {n_points} up to {n_points+11}, giving the predicted values.

# # # # # # # 5. **Floor at zero** — any predicted value below 0 is clamped to 0.

# # # # # # # ---

# # # # # # # **Trend direction: {direction}**
# # # # # # # Each month the value changes by approximately **{slope:+.3f} {units_chosen}**.

# # # # # # # *Note: This is a simple linear extrapolation based on the available baseline months.
# # # # # # # It works well for stable metrics. It does not account for seasonality or sudden
# # # # # # # operational changes.*
# # # # # # # """)
# # # # # # #     else:
# # # # # # #         st.warning("Not enough non-zero baseline data to build a trend.")

# # # # # # # # ── Download ──────────────────────────────────────────────────────────────────
# # # # # # # st.download_button(
# # # # # # #     "⬇ Download full data (CSV)",
# # # # # # #     df.to_csv(index=False).encode("utf-8"),
# # # # # # #     file_name="ghg_summary.csv",
# # # # # # #     mime="text/csv"
# # # # # # # )





# # # # # # # //new code


# # # # # # import streamlit as st
# # # # # # import pandas as pd
# # # # # # import numpy as np
# # # # # # import plotly.express as px
# # # # # # import re

# # # # # # st.set_page_config(page_title="GHG Summary Analyzer", layout="wide")
# # # # # # st.title("GHG Summary Analyzer")
# # # # # # st.caption("Upload an Excel; we’ll detect the Summary sheet, clean it, and build interactive charts.")

# # # # # # uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])


# # # # # # # ---------------- Helpers ----------------
# # # # # # def load_excel(uploaded_file):
# # # # # #     for eng in ("openpyxl", "calamine"):
# # # # # #         try:
# # # # # #             return pd.ExcelFile(uploaded_file, engine=eng), eng
# # # # # #         except Exception:
# # # # # #             continue
# # # # # #     raise ImportError(
# # # # # #         "No Excel engine available. Install one of:\n"
# # # # # #         "  pip3 install openpyxl\n"
# # # # # #         "  or\n"
# # # # # #         "  pip3 install pandas-calamine"
# # # # # #     )

# # # # # # def find_sheet(xls):
# # # # # #     if "Summary Sheet" in xls.sheet_names:
# # # # # #         return "Summary Sheet"
# # # # # #     for s in xls.sheet_names:
# # # # # #         if "summary" in s.lower():
# # # # # #             return s
# # # # # #     return xls.sheet_names[0]

# # # # # # def header_detect_clean(df_raw: pd.DataFrame) -> pd.DataFrame:
# # # # # #     raw = df_raw.dropna(how="all").dropna(axis=1, how="all").copy()
# # # # # #     header_idx = raw.head(min(15, len(raw))).notna().sum(axis=1).idxmax()
# # # # # #     df = raw.copy()
# # # # # #     df.columns = df.loc[header_idx].astype(str).str.strip()
# # # # # #     df = df.loc[header_idx + 1 :].reset_index(drop=True)

# # # # # #     # make unique column names
# # # # # #     seen, cols = {}, []
# # # # # #     for c in df.columns:
# # # # # #         base = c if c and c != "nan" else "Unnamed"
# # # # # #         seen[base] = seen.get(base, 0) + 1
# # # # # #         cols.append(base if seen[base] == 1 else f"{base}_{seen[base]}")
# # # # # #     df.columns = cols

# # # # # #     # numeric coercion when >=60% convertible
# # # # # #     for c in df.columns:
# # # # # #         parsed = pd.to_numeric(df[c], errors="coerce")
# # # # # #         if parsed.notna().mean() >= 0.6:
# # # # # #             df[c] = parsed
# # # # # #     return df

# # # # # # def prune_columns(df: pd.DataFrame, drop_unnamed=True, null_threshold=0.90, drop_pattern=None):
# # # # # #     drop = []
# # # # # #     for c in df.columns:
# # # # # #         name = str(c).strip()
# # # # # #         if drop_unnamed and (name.lower().startswith("unnamed") or name == "" or name.lower() == "nan"):
# # # # # #             drop.append(c)
# # # # # #             continue
# # # # # #         if drop_pattern and re.match(drop_pattern, name):
# # # # # #             drop.append(c)
# # # # # #             continue
# # # # # #         if df[c].isna().mean() >= null_threshold:
# # # # # #             drop.append(c)
# # # # # #             continue
# # # # # #         if df[c].dtype == "object":
# # # # # #             s = df[c].astype(str).str.strip().replace("nan", "")
# # # # # #             if s.replace("", np.nan).dropna().empty:
# # # # # #                 drop.append(c)
# # # # # #                 continue
# # # # # #     df_clean = df.drop(columns=drop)
# # # # # #     return df_clean, drop

# # # # # # def guess_label_cols(df, num_cols):
# # # # # #     return [c for c in df.columns if c not in num_cols]

# # # # # # def arrow_safe(df: pd.DataFrame) -> pd.DataFrame:
# # # # # #     out = df.copy()
# # # # # #     for c in out.columns:
# # # # # #         if out[c].dtype == "object":
# # # # # #             out[c] = out[c].astype("string")
# # # # # #     return out

# # # # # # def reduce_ticks(labels, max_ticks=25):
# # # # # #     n = len(labels)
# # # # # #     if n <= max_ticks:
# # # # # #         return list(range(n)), labels
# # # # # #     step = max(1, n // max_ticks)
# # # # # #     idxs = list(range(0, n, step))
# # # # # #     return idxs, [labels[i] for i in idxs]

# # # # # # def hex_to_rgba(hex_color: str, alpha: float) -> str:
# # # # # #     """Convert #RRGGBB to rgba(...) string (alpha 0..1).
# # # # # #        If the color is not a hex string, return the original color (Plotly will use it as-is)."""
# # # # # #     try:
# # # # # #         hc = str(hex_color).lstrip("#")
# # # # # #         if len(hc) != 6:
# # # # # #             return hex_color
# # # # # #         r = int(hc[0:2], 16)
# # # # # #         g = int(hc[2:4], 16)
# # # # # #         b = int(hc[4:6], 16)
# # # # # #         return f"rgba({r},{g},{b},{alpha})"
# # # # # #     except Exception:
# # # # # #         return hex_color

# # # # # # def palette_color(palette: list, idx: int) -> str:
# # # # # #     return palette[idx % len(palette)]

# # # # # # PALETTES = {
# # # # # #     "Plotly": px.colors.qualitative.Plotly,
# # # # # #     "D3": px.colors.qualitative.D3,
# # # # # #     "Bold": px.colors.qualitative.Bold,
# # # # # #     "Dark24": px.colors.qualitative.Dark24,
# # # # # #     "G10": px.colors.qualitative.G10,
# # # # # #     "Alphabet": px.colors.qualitative.Alphabet,
# # # # # # }

# # # # # # # ---------------- App ----------------
# # # # # # if uploaded:
# # # # # #     try:
# # # # # #         xls, engine_used = load_excel(uploaded)
# # # # # #         st.caption(f"Excel engine: **{engine_used}**")
# # # # # #         sheet = find_sheet(xls)
# # # # # #         st.success(f"Using sheet: {sheet}")
# # # # # #         raw = pd.read_excel(uploaded, sheet_name=sheet, engine=engine_used, header=None)
# # # # # #         df = header_detect_clean(raw)
# # # # # #     except Exception as e:
# # # # # #         st.exception(e)
# # # # # #         st.stop()

    
# # # # # #     drop_pattern = r"^0\.0_.*"
# # # # # #     df, dropped_prune = prune_columns(df, drop_unnamed=True, null_threshold=0.90, drop_pattern=drop_pattern)
# # # # # #     if dropped_prune:
# # # # # #         st.info(f"Hidden {len(dropped_prune)} mostly-empty / unnamed / patterned columns.")

# # # # # #     # detect numeric columns after pruning
# # # # # #     num_cols = df.select_dtypes(include="number").columns.tolist()

# # # # # #     # filter numeric columns to exclude any that are all-NaN OR all-zero
# # # # # #     valid_y_cols = []
# # # # # #     excluded_y_cols = []
# # # # # #     for c in num_cols:
# # # # # #         s = pd.to_numeric(df[c], errors="coerce")
# # # # # #         s_non_na = s.dropna()
# # # # # #         if s_non_na.empty:
# # # # # #             excluded_y_cols.append(c)
# # # # # #         elif s_non_na.abs().sum() == 0:
# # # # # #             excluded_y_cols.append(c)
# # # # # #         else:
# # # # # #             valid_y_cols.append(c)

# # # # # #     # label/categorical candidates
# # # # # #     cat_cols = guess_label_cols(df, num_cols)

# # # # # #     # ---------------- SIDEBAR ----------------
# # # # # #     with st.sidebar:
# # # # # #         st.header("Controls")

# # # # # #         # X-axis chooser
# # # # # #         x_col = st.selectbox("X-axis (categorical / time)", options=(cat_cols or [None]))

# # # # # #         # Y-axis chooser: only show numeric columns that have at least one non-zero value
# # # # # #         if valid_y_cols:
# # # # # #             y_cols = st.multiselect("Y-axis (numeric)", options=valid_y_cols, default=valid_y_cols[:1])
# # # # # #         else:
# # # # # #             st.warning("No numeric columns with non-zero/nonnull values were detected.")
# # # # # #             y_cols = []

    
# # # # # #         if y_cols:
# # # # # #             y_numeric_all = df[y_cols].apply(pd.to_numeric, errors="coerce").fillna(0).abs()
# # # # # #             nonzero_mask_for_filters = (y_numeric_all.sum(axis=1) != 0)
# # # # # #             df_for_filters = df.loc[nonzero_mask_for_filters].copy()
# # # # # #         else:
# # # # # #             df_for_filters = df.copy()

# # # # # #         st.markdown("**Filters**")
# # # # # #         active_filters = {}
# # # # # #         for c in cat_cols[:6]:
# # # # # #             vals = df_for_filters[c].dropna().astype(str).unique().tolist()
# # # # # #             if 1 < len(vals) <= 200:
# # # # # #                 sel = st.multiselect(f"{c}", options=vals, default=[], placeholder="(all)")
# # # # # #                 if sel:
# # # # # #                     active_filters[c] = set(sel)

# # # # # #         chart_type = st.radio("Chart type", ["Line", "Bar", "Area"], index=1)
# # # # # #         # use_log = st.checkbox("Log scale (Y)", value=False)
        

# # # # # #     # ---------------- APPLY FILTERS & AGGREGATION ----------------
# # # # # #     filt = pd.Series(True, index=df.index)
# # # # # #     for col, allowed in active_filters.items():
# # # # # #         filt &= df[col].astype(str).isin(allowed)
# # # # # #     df_f = df.loc[filt].copy()

  

# # # # # #     if y_cols:
# # # # # #         y_numeric_after = df_f[y_cols].apply(pd.to_numeric, errors="coerce")
# # # # # #         nonzero_mask_after = (y_numeric_after.fillna(0).abs().sum(axis=1) != 0)
# # # # # #         removed_count = (~nonzero_mask_after).sum()
# # # # # #         if removed_count:
# # # # # #             st.info(f"Removed {int(removed_count)} row(s) where selected Y columns were all zero/null.")
# # # # # #         df_f = df_f.loc[nonzero_mask_after].reset_index(drop=True)

# # # # # #     if excluded_y_cols:
# # # # # #         st.info(f"Excluded numeric columns (all zero or all null): {', '.join(excluded_y_cols)}")

# # # # # #     if df_f.empty:
# # # # # #         st.warning("No rows left after applying filters / removing zero rows. Adjust filters or Y selection.")
# # # # # #         st.stop()

# # # # # #     # Table
# # # # # #     st.subheader("Table")
# # # # # #     st.dataframe(arrow_safe(df_f), width="stretch", height=420)

# # # # # #     if not y_cols:
# # # # # #         st.warning("Pick at least one numeric Y column.")
# # # # # #         st.stop()

# # # # # #     if x_col and x_col in df_f.columns:
# # # # # #         x_vals = df_f[x_col].astype(str).fillna("")
# # # # # #     else:
# # # # # #         x_vals = pd.Series(range(len(df_f)), index=df_f.index, name="Index")
# # # # # #         x_col = "Index"
# # # # # #         df_f = df_f.assign(Index=x_vals)

   
# # # # # #     for i, y in enumerate(y_cols):
# # # # # #         primary_color = px.colors.qualitative.Plotly[i % len(px.colors.qualitative.Plotly)]
# # # # # #         y_vals = pd.to_numeric(df_f[y], errors="coerce")
# # # # # #         if y_vals.dropna().empty or y_vals.dropna().abs().sum() == 0:
# # # # # #             st.info(f"Skipping chart for '{y}' (all values are zero or null after filters).")
# # # # # #             continue

        

# # # # # #         y_vals_filled = y_vals.fillna(0)
# # # # # #         df_plot = pd.DataFrame({x_col: x_vals, y: y_vals_filled})

# # # # # #         if chart_type == "Bar":
# # # # # #             fig = px.bar(df_plot, x=x_col, y=y)
# # # # # #             # bar traces use marker.color or marker=dict(color=...)
# # # # # #             fig.update_traces(marker=dict(color=primary_color))
# # # # # #         elif chart_type == "Line":
# # # # # #             fig = px.line(df_plot, x=x_col, y=y)
# # # # # #             # line traces accept line.color
# # # # # #             fig.update_traces(line=dict(color=primary_color))
# # # # # #         else:  # Area
# # # # # #             fig = px.area(df_plot, x=x_col, y=y)
# # # # # #             # area is a scatter with fill; set line color and a semi-transparent fillcolor
# # # # # #             fig.update_traces(line=dict(color=primary_color),
# # # # # #                               fillcolor=hex_to_rgba(primary_color, 0.25))

        

     

# # # # # #         fig.update_layout(
         
# # # # # #             margin=dict(l=10, r=10, t=40, b=10),
# # # # # #             height=480,
# # # # # #             hovermode="x unified",
# # # # # #         )

# # # # # #         st.plotly_chart(fig, use_container_width=True)

   
# # # # # #     st.download_button(
# # # # # #         "Download cleaned/filtered data (CSV)",
# # # # # #         df_f.to_csv(index=False).encode("utf-8"),
# # # # # #         file_name="summary_cleaned_filtered.csv",
# # # # # #         mime="text/csv",
# # # # # #     )

# # # # # # else:
# # # # # #     st.info("Upload a file to begin.")



# # # # # import streamlit as st
# # # # # import pandas as pd
# # # # # import numpy as np
# # # # # import plotly.express as px
# # # # # import plotly.graph_objects as go
# # # # # import re
# # # # # from scipy import stats

# # # # # st.set_page_config(page_title="GHG Summary Analyzer", layout="wide")
# # # # # st.title("GHG Summary Analyzer")
# # # # # st.caption("Upload an Excel; we'll detect the Summary sheet, clean it, and build interactive charts.")

# # # # # uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])


# # # # # # ---------------- Helpers ----------------
# # # # # def load_excel(uploaded_file):
# # # # #     for eng in ("openpyxl", "calamine"):
# # # # #         try:
# # # # #             return pd.ExcelFile(uploaded_file, engine=eng), eng
# # # # #         except Exception:
# # # # #             continue
# # # # #     raise ImportError(
# # # # #         "No Excel engine available. Install one of:\n"
# # # # #         "  pip3 install openpyxl\n"
# # # # #         "  or\n"
# # # # #         "  pip3 install pandas-calamine"
# # # # #     )

# # # # # def find_sheet(xls):
# # # # #     if "Summary Sheet" in xls.sheet_names:
# # # # #         return "Summary Sheet"
# # # # #     for s in xls.sheet_names:
# # # # #         if "summary" in s.lower():
# # # # #             return s
# # # # #     return xls.sheet_names[0]

# # # # # def header_detect_clean(df_raw: pd.DataFrame) -> pd.DataFrame:
# # # # #     raw = df_raw.dropna(how="all").dropna(axis=1, how="all").copy()
# # # # #     header_idx = raw.head(min(15, len(raw))).notna().sum(axis=1).idxmax()
# # # # #     df = raw.copy()
# # # # #     df.columns = df.loc[header_idx].astype(str).str.strip()
# # # # #     df = df.loc[header_idx + 1 :].reset_index(drop=True)

# # # # #     seen, cols = {}, []
# # # # #     for c in df.columns:
# # # # #         base = c if c and c != "nan" else "Unnamed"
# # # # #         seen[base] = seen.get(base, 0) + 1
# # # # #         cols.append(base if seen[base] == 1 else f"{base}_{seen[base]}")
# # # # #     df.columns = cols

# # # # #     for c in df.columns:
# # # # #         parsed = pd.to_numeric(df[c], errors="coerce")
# # # # #         if parsed.notna().mean() >= 0.6:
# # # # #             df[c] = parsed
# # # # #     return df

# # # # # def prune_columns(df: pd.DataFrame, drop_unnamed=True, null_threshold=0.90, drop_pattern=None):
# # # # #     drop = []
# # # # #     for c in df.columns:
# # # # #         name = str(c).strip()
# # # # #         if drop_unnamed and (name.lower().startswith("unnamed") or name == "" or name.lower() == "nan"):
# # # # #             drop.append(c)
# # # # #             continue
# # # # #         if drop_pattern and re.match(drop_pattern, name):
# # # # #             drop.append(c)
# # # # #             continue
# # # # #         if df[c].isna().mean() >= null_threshold:
# # # # #             drop.append(c)
# # # # #             continue
# # # # #         if df[c].dtype == "object":
# # # # #             s = df[c].astype(str).str.strip().replace("nan", "")
# # # # #             if s.replace("", np.nan).dropna().empty:
# # # # #                 drop.append(c)
# # # # #                 continue
# # # # #     df_clean = df.drop(columns=drop)
# # # # #     return df_clean, drop

# # # # # def guess_label_cols(df, num_cols):
# # # # #     return [c for c in df.columns if c not in num_cols]

# # # # # def arrow_safe(df: pd.DataFrame) -> pd.DataFrame:
# # # # #     out = df.copy()
# # # # #     for c in out.columns:
# # # # #         if out[c].dtype == "object":
# # # # #             out[c] = out[c].astype("string")
# # # # #     return out

# # # # # def hex_to_rgba(hex_color: str, alpha: float) -> str:
# # # # #     try:
# # # # #         hc = str(hex_color).lstrip("#")
# # # # #         if len(hc) != 6:
# # # # #             return hex_color
# # # # #         r = int(hc[0:2], 16)
# # # # #         g = int(hc[2:4], 16)
# # # # #         b = int(hc[4:6], 16)
# # # # #         return f"rgba({r},{g},{b},{alpha})"
# # # # #     except Exception:
# # # # #         return hex_color

# # # # # def palette_color(palette: list, idx: int) -> str:
# # # # #     return palette[idx % len(palette)]

# # # # # PALETTES = {
# # # # #     "Plotly": px.colors.qualitative.Plotly,
# # # # #     "D3": px.colors.qualitative.D3,
# # # # #     "Bold": px.colors.qualitative.Bold,
# # # # #     "Dark24": px.colors.qualitative.Dark24,
# # # # #     "G10": px.colors.qualitative.G10,
# # # # #     "Alphabet": px.colors.qualitative.Alphabet,
# # # # # }

# # # # # # ─────────────────────────────────────────────────────────────────
# # # # # # GHG DEDICATED PARSER — reads raw sheet and extracts GHG rows
# # # # # # ─────────────────────────────────────────────────────────────────
# # # # # GHG_ROW_LABELS = {
# # # # #     "Emission per ton of Equivalent product":       "T",
# # # # #     "Total Direct and Indirect Emission":            "S",
# # # # #     "Total Direct Emission (Scope 1)":               "Scope1",
# # # # #     "Total Indirect Emission (Scope 2)":             "Scope2",
# # # # #     "Total Equivalent Product for GHG Emission Intensity": "P",
# # # # # }

# # # # # KNOWN_MONTHS = ["January","February","March","April","May","June",
# # # # #                 "July","August","September","October","November","December"]

# # # # # def extract_ghg_data(raw: pd.DataFrame):
# # # # #     """
# # # # #     Scan every cell for known GHG row labels; extract monthly numeric values.
# # # # #     Returns a dict: label -> list of (month, value)
# # # # #     Also returns the list of months found.
# # # # #     """
# # # # #     # Find month header row
# # # # #     months_found = []
# # # # #     month_header_row = None
# # # # #     for i, row in raw.iterrows():
# # # # #         row_str = row.astype(str).str.strip().tolist()
# # # # #         hits = [c for c in row_str if any(m.lower() in c.lower() for m in KNOWN_MONTHS)]
# # # # #         if len(hits) >= 2:
# # # # #             months_found = [c for c in row_str if any(m.lower() in c.lower() for m in KNOWN_MONTHS)]
# # # # #             # keep order
# # # # #             month_header_row = i
# # # # #             break

# # # # #     # Fallback: use column headers if months found there
# # # # #     col_months = []
# # # # #     for c in raw.columns:
# # # # #         cs = str(c).strip()
# # # # #         if any(m.lower() in cs.lower() for m in KNOWN_MONTHS):
# # # # #             col_months.append(cs)

# # # # #     # Find which raw columns hold monthly data
# # # # #     # Strategy: look at row 4 (S. No. / Particulars / Unit / April / May / June / ...)
# # # # #     # We'll scan all rows for the month names
# # # # #     month_cols = {}  # month_name -> col index
# # # # #     for i, row in raw.iterrows():
# # # # #         for col_idx, val in enumerate(row):
# # # # #             vs = str(val).strip()
# # # # #             for m in KNOWN_MONTHS:
# # # # #                 if m.lower() == vs.lower():
# # # # #                     month_cols[m] = col_idx
# # # # #         if len(month_cols) >= 2:
# # # # #             break

# # # # #     results = {label: {} for label in GHG_ROW_LABELS}

# # # # #     for i, row in raw.iterrows():
# # # # #         row_vals = [str(v).strip() for v in row]
# # # # #         for label in GHG_ROW_LABELS:
# # # # #             if any(label.lower() in rv.lower() for rv in row_vals):
# # # # #                 for month, col_idx in month_cols.items():
# # # # #                     try:
# # # # #                         v = pd.to_numeric(row.iloc[col_idx], errors="coerce")
# # # # #                         if pd.notna(v):
# # # # #                             results[label][month] = v
# # # # #                     except Exception:
# # # # #                         pass

# # # # #     ordered_months = [m for m in KNOWN_MONTHS if m in month_cols]
# # # # #     return results, ordered_months


# # # # # def predict_remaining_months(months_known: list, values: list, all_months: list):
# # # # #     """
# # # # #     Given known (month, value) pairs, predict remaining months via linear regression.
# # # # #     months_known: list of str, e.g. ['April','May','June']
# # # # #     values: corresponding floats
# # # # #     all_months: full 12-month list
# # # # #     Returns dict: month -> predicted_value for unknown months
# # # # #     """
# # # # #     if len(values) < 2:
# # # # #         return {}

# # # # #     x_known = np.array([all_months.index(m) for m in months_known if m in all_months])
# # # # #     y_known = np.array(values)

# # # # #     # Linear regression
# # # # #     slope, intercept, r, p, se = stats.linregress(x_known, y_known)

# # # # #     predictions = {}
# # # # #     for m in all_months:
# # # # #         if m not in months_known:
# # # # #             idx = all_months.index(m)
# # # # #             predictions[m] = slope * idx + intercept
# # # # #     return predictions, slope, intercept


# # # # # def make_ghg_chart(
# # # # #     ghg_data: dict,
# # # # #     label: str,
# # # # #     months_known: list,
# # # # #     all_months: list,
# # # # #     chart_type: str = "Bar",
# # # # #     add_target_line: bool = False,
# # # # #     target_value: float = 0.3552,
# # # # #     add_prediction: bool = True,
# # # # #     y_axis_label: str = "GEI",
# # # # #     color: str = "#2196F3",
# # # # # ):
# # # # #     values_known = [ghg_data[label].get(m) for m in months_known]
# # # # #     values_known_clean = [v for v in values_known if v is not None]

# # # # #     # Predict
# # # # #     pred_dict = {}
# # # # #     if add_prediction and len(values_known_clean) >= 2:
# # # # #         pred_result = predict_remaining_months(
# # # # #             [m for m, v in zip(months_known, values_known) if v is not None],
# # # # #             values_known_clean,
# # # # #             all_months,
# # # # #         )
# # # # #         if isinstance(pred_result, tuple):
# # # # #             pred_dict, slope, intercept = pred_result

# # # # #     # Build full month x-axis
# # # # #     all_x = all_months
# # # # #     actual_y = [ghg_data[label].get(m, None) for m in all_months]
# # # # #     pred_y = [pred_dict.get(m, None) for m in all_months]

# # # # #     fig = go.Figure()

# # # # #     # ── Actual data trace ──────────────────────────────────────────
# # # # #     rgba_fill = hex_to_rgba(color, 0.25)
# # # # #     if chart_type == "Bar":
# # # # #         fig.add_trace(go.Bar(
# # # # #             x=all_x,
# # # # #             y=[v if v is not None else 0 for v in actual_y],
# # # # #             name="Actual",
# # # # #             marker_color=[color if v is not None else "rgba(0,0,0,0)" for v in actual_y],
# # # # #             text=[f"{v:.4f}" if v is not None else "" for v in actual_y],
# # # # #             textposition="outside",
# # # # #         ))
# # # # #     elif chart_type == "Line":
# # # # #         fig.add_trace(go.Scatter(
# # # # #             x=all_x, y=actual_y,
# # # # #             name="Actual",
# # # # #             mode="lines+markers+text",
# # # # #             line=dict(color=color, width=2.5),
# # # # #             marker=dict(size=8),
# # # # #             text=[f"{v:.4f}" if v is not None else "" for v in actual_y],
# # # # #             textposition="top center",
# # # # #         ))
# # # # #     else:  # Area
# # # # #         fig.add_trace(go.Scatter(
# # # # #             x=all_x, y=actual_y,
# # # # #             name="Actual",
# # # # #             mode="lines+markers",
# # # # #             fill="tozeroy",
# # # # #             line=dict(color=color, width=2.5),
# # # # #             fillcolor=rgba_fill,
# # # # #             marker=dict(size=8),
# # # # #         ))

# # # # #     # ── Predicted months (dashed) ──────────────────────────────────
# # # # #     if pred_dict:
# # # # #         pred_x = [m for m in all_months if m in pred_dict]
# # # # #         pred_vals = [pred_dict[m] for m in pred_x]
# # # # #         fig.add_trace(go.Scatter(
# # # # #             x=pred_x, y=pred_vals,
# # # # #             name="Predicted (Baseline)",
# # # # #             mode="lines+markers+text",
# # # # #             line=dict(color="#FF9800", width=2, dash="dash"),
# # # # #             marker=dict(size=8, symbol="diamond"),
# # # # #             text=[f"{v:.4f}" for v in pred_vals],
# # # # #             textposition="top center",
# # # # #         ))

# # # # #     # ── Static target / baseline red line ─────────────────────────
# # # # #     if add_target_line:
# # # # #         fig.add_hline(
# # # # #             y=target_value,
# # # # #             line=dict(color="red", width=2, dash="dot"),
# # # # #             annotation_text=f"Baseline Target: {target_value}",
# # # # #             annotation_position="top right",
# # # # #             annotation_font=dict(color="red", size=12),
# # # # #         )

# # # # #     fig.update_layout(
# # # # #         title=dict(text=label, font=dict(size=15)),
# # # # #         xaxis_title="Month",
# # # # #         yaxis_title=y_axis_label,
# # # # #         height=480,
# # # # #         margin=dict(l=10, r=10, t=50, b=10),
# # # # #         hovermode="x unified",
# # # # #         legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
# # # # #     )
# # # # #     return fig


# # # # # # ─────────────────────────────────────────────────────────────────
# # # # # # APP
# # # # # # ─────────────────────────────────────────────────────────────────
# # # # # if uploaded:
# # # # #     try:
# # # # #         xls, engine_used = load_excel(uploaded)
# # # # #         st.caption(f"Excel engine: **{engine_used}**")
# # # # #         sheet = find_sheet(xls)
# # # # #         st.success(f"Using sheet: **{sheet}**")
# # # # #         raw = pd.read_excel(uploaded, sheet_name=sheet, engine=engine_used, header=None)
# # # # #         df = header_detect_clean(raw)
# # # # #     except Exception as e:
# # # # #         st.exception(e)
# # # # #         st.stop()

# # # # #     drop_pattern = r"^0\.0_.*"
# # # # #     df, dropped_prune = prune_columns(df, drop_unnamed=True, null_threshold=0.90, drop_pattern=drop_pattern)
# # # # #     if dropped_prune:
# # # # #         st.info(f"Hidden {len(dropped_prune)} mostly-empty / unnamed / patterned columns.")

# # # # #     num_cols = df.select_dtypes(include="number").columns.tolist()
# # # # #     valid_y_cols = []
# # # # #     excluded_y_cols = []
# # # # #     for c in num_cols:
# # # # #         s = pd.to_numeric(df[c], errors="coerce")
# # # # #         s_non_na = s.dropna()
# # # # #         if s_non_na.empty or s_non_na.abs().sum() == 0:
# # # # #             excluded_y_cols.append(c)
# # # # #         else:
# # # # #             valid_y_cols.append(c)

# # # # #     cat_cols = guess_label_cols(df, num_cols)

# # # # #     # ── Extract GHG rows from raw sheet ────────────────────────────
# # # # #     ghg_data, months_known = extract_ghg_data(raw)

# # # # #     # ─────────────────────────────────────────────────────────────
# # # # #     # SIDEBAR
# # # # #     # ─────────────────────────────────────────────────────────────
# # # # #     with st.sidebar:
# # # # #         st.header("Controls")
# # # # #         x_col = st.selectbox("X-axis (categorical / time)", options=(cat_cols or [None]))
# # # # #         if valid_y_cols:
# # # # #             y_cols = st.multiselect("Y-axis (numeric)", options=valid_y_cols, default=valid_y_cols[:1])
# # # # #         else:
# # # # #             st.warning("No numeric columns with non-zero/nonnull values were detected.")
# # # # #             y_cols = []

# # # # #         if y_cols:
# # # # #             y_numeric_all = df[y_cols].apply(pd.to_numeric, errors="coerce").fillna(0).abs()
# # # # #             nonzero_mask_for_filters = (y_numeric_all.sum(axis=1) != 0)
# # # # #             df_for_filters = df.loc[nonzero_mask_for_filters].copy()
# # # # #         else:
# # # # #             df_for_filters = df.copy()

# # # # #         st.markdown("**Filters**")
# # # # #         active_filters = {}
# # # # #         for c in cat_cols[:6]:
# # # # #             vals = df_for_filters[c].dropna().astype(str).unique().tolist()
# # # # #             if 1 < len(vals) <= 200:
# # # # #                 sel = st.multiselect(f"{c}", options=vals, default=[], placeholder="(all)")
# # # # #                 if sel:
# # # # #                     active_filters[c] = set(sel)

# # # # #         chart_type = st.radio("Chart type", ["Line", "Bar", "Area"], index=1)

# # # # #         st.markdown("---")
# # # # #         st.subheader("GHG Chart Options")
# # # # #         ghg_chart_type = st.radio("GHG Chart type", ["Line", "Bar", "Area"], index=0)
# # # # #         show_prediction = st.checkbox("Show Predicted Months", value=True)
# # # # #         baseline_value = st.number_input(
# # # # #             "Baseline Target (red line)",
# # # # #             value=0.3552, step=0.001, format="%.4f",
# # # # #             help="Static red reference line shown on Emission Intensity chart"
# # # # #         )

# # # # #     # ─────────────────────────────────────────────────────────────
# # # # #     # FILTERS
# # # # #     # ─────────────────────────────────────────────────────────────
# # # # #     filt = pd.Series(True, index=df.index)
# # # # #     for col, allowed in active_filters.items():
# # # # #         filt &= df[col].astype(str).isin(allowed)
# # # # #     df_f = df.loc[filt].copy()

# # # # #     if y_cols:
# # # # #         y_numeric_after = df_f[y_cols].apply(pd.to_numeric, errors="coerce")
# # # # #         nonzero_mask_after = (y_numeric_after.fillna(0).abs().sum(axis=1) != 0)
# # # # #         removed_count = (~nonzero_mask_after).sum()
# # # # #         if removed_count:
# # # # #             st.info(f"Removed {int(removed_count)} row(s) where selected Y columns were all zero/null.")
# # # # #         df_f = df_f.loc[nonzero_mask_after].reset_index(drop=True)

# # # # #     if excluded_y_cols:
# # # # #         st.info(f"Excluded numeric columns (all zero or all null): {', '.join(excluded_y_cols)}")

# # # # #     if df_f.empty:
# # # # #         st.warning("No rows left after applying filters / removing zero rows. Adjust filters or Y selection.")
# # # # #         st.stop()

# # # # #     # Table
# # # # #     st.subheader("Data Table")
# # # # #     st.dataframe(arrow_safe(df_f), use_container_width=True, height=420)

# # # # #     # ─────────────────────────────────────────────────────────────
# # # # #     # ★ GHG DEDICATED SECTION ★
# # # # #     # ─────────────────────────────────────────────────────────────
# # # # #     any_ghg = any(bool(ghg_data[lbl]) for lbl in GHG_ROW_LABELS)

# # # # #     if any_ghg and months_known:
# # # # #         st.markdown("---")
# # # # #         st.header("📊 GHG Emission Analysis")

# # # # #         col_info = st.columns(len(months_known))
# # # # #         for i, m in enumerate(months_known):
# # # # #             v = ghg_data["Emission per ton of Equivalent product"].get(m)
# # # # #             col_info[i].metric(
# # # # #                 label=m,
# # # # #                 value=f"{v:.4f}" if v is not None else "N/A",
# # # # #                 delta=f"vs target {baseline_value:.4f}" if v is not None else None,
# # # # #                 delta_color="inverse",
# # # # #             )

# # # # #         CHART_CONFIGS = [
# # # # #             {
# # # # #                 "label": "Emission per ton of Equivalent product",
# # # # #                 "y_axis": "tCO2 eq/ton",
# # # # #                 "color": "#1976D2",
# # # # #                 "add_target": True,
# # # # #                 "title_extra": " (with Baseline Target & Forecast)",
# # # # #             },
# # # # #             {
# # # # #                 "label": "Total Direct and Indirect Emission",
# # # # #                 "y_axis": "tCO2 eq",
# # # # #                 "color": "#E53935",
# # # # #                 "add_target": False,
# # # # #                 "title_extra": "",
# # # # #             },
# # # # #             {
# # # # #                 "label": "Total Direct Emission (Scope 1)",
# # # # #                 "y_axis": "tCO2 eq",
# # # # #                 "color": "#FF7043",
# # # # #                 "add_target": False,
# # # # #                 "title_extra": "",
# # # # #             },
# # # # #             {
# # # # #                 "label": "Total Indirect Emission (Scope 2)",
# # # # #                 "y_axis": "tCO2 eq",
# # # # #                 "color": "#7B1FA2",
# # # # #                 "add_target": False,
# # # # #                 "title_extra": "",
# # # # #             },
# # # # #             {
# # # # #                 "label": "Total Equivalent Product for GHG Emission Intensity",
# # # # #                 "y_axis": "Tonnes",
# # # # #                 "color": "#2E7D32",
# # # # #                 "add_target": False,
# # # # #                 "title_extra": "",
# # # # #             },
# # # # #         ]

# # # # #         for cfg in CHART_CONFIGS:
# # # # #             lbl = cfg["label"]
# # # # #             if not ghg_data.get(lbl):
# # # # #                 st.warning(f"No data found for: {lbl}")
# # # # #                 continue

# # # # #             fig = make_ghg_chart(
# # # # #                 ghg_data=ghg_data,
# # # # #                 label=lbl,
# # # # #                 months_known=months_known,
# # # # #                 all_months=KNOWN_MONTHS,
# # # # #                 chart_type=ghg_chart_type,
# # # # #                 add_target_line=cfg["add_target"],
# # # # #                 target_value=baseline_value,
# # # # #                 add_prediction=show_prediction,
# # # # #                 y_axis_label=cfg["y_axis"],
# # # # #                 color=cfg["color"],
# # # # #             )
# # # # #             st.subheader(lbl + cfg.get("title_extra", ""))
# # # # #             st.plotly_chart(fig, use_container_width=True)

# # # # #         # Prediction summary table
# # # # #         if show_prediction:
# # # # #             st.subheader("📅 Baseline Prediction — Remaining Months")
# # # # #             pred_rows = []
# # # # #             for lbl in GHG_ROW_LABELS:
# # # # #                 vals = ghg_data[lbl]
# # # # #                 if len(vals) < 2:
# # # # #                     continue
# # # # #                 mk = [m for m in KNOWN_MONTHS if m in vals]
# # # # #                 mv = [vals[m] for m in mk]
# # # # #                 try:
# # # # #                     pred_dict, slope, intercept = predict_remaining_months(mk, mv, KNOWN_MONTHS)
# # # # #                     for m, pv in pred_dict.items():
# # # # #                         pred_rows.append({
# # # # #                             "Metric": lbl,
# # # # #                             "Month": m,
# # # # #                             "Predicted Value": round(pv, 6),
# # # # #                         })
# # # # #                 except Exception:
# # # # #                     pass

# # # # #             if pred_rows:
# # # # #                 pred_df = pd.DataFrame(pred_rows)
# # # # #                 st.dataframe(pred_df, use_container_width=True)

# # # # #     # ─────────────────────────────────────────────────────────────
# # # # #     # ORIGINAL GENERIC CHARTS
# # # # #     # ─────────────────────────────────────────────────────────────
# # # # #     if not y_cols:
# # # # #         st.warning("Pick at least one numeric Y column.")
# # # # #         st.stop()

# # # # #     st.markdown("---")
# # # # #     st.header("📈 Generic Charts (Selected Columns)")

# # # # #     if x_col and x_col in df_f.columns:
# # # # #         x_vals = df_f[x_col].astype(str).fillna("")
# # # # #     else:
# # # # #         x_vals = pd.Series(range(len(df_f)), index=df_f.index, name="Index")
# # # # #         x_col = "Index"
# # # # #         df_f = df_f.assign(Index=x_vals)

# # # # #     for i, y in enumerate(y_cols):
# # # # #         primary_color = px.colors.qualitative.Plotly[i % len(px.colors.qualitative.Plotly)]
# # # # #         y_vals = pd.to_numeric(df_f[y], errors="coerce")
# # # # #         if y_vals.dropna().empty or y_vals.dropna().abs().sum() == 0:
# # # # #             st.info(f"Skipping chart for '{y}' (all values are zero or null after filters).")
# # # # #             continue

# # # # #         y_vals_filled = y_vals.fillna(0)
# # # # #         df_plot = pd.DataFrame({x_col: x_vals, y: y_vals_filled})

# # # # #         if chart_type == "Bar":
# # # # #             fig = px.bar(df_plot, x=x_col, y=y)
# # # # #             fig.update_traces(marker=dict(color=primary_color))
# # # # #         elif chart_type == "Line":
# # # # #             fig = px.line(df_plot, x=x_col, y=y)
# # # # #             fig.update_traces(line=dict(color=primary_color))
# # # # #         else:
# # # # #             fig = px.area(df_plot, x=x_col, y=y)
# # # # #             fig.update_traces(line=dict(color=primary_color),
# # # # #                               fillcolor=hex_to_rgba(primary_color, 0.25))

# # # # #         fig.update_layout(
# # # # #             margin=dict(l=10, r=10, t=40, b=10),
# # # # #             height=480,
# # # # #             hovermode="x unified",
# # # # #         )
# # # # #         st.plotly_chart(fig, use_container_width=True)

# # # # #     st.download_button(
# # # # #         "Download cleaned/filtered data (CSV)",
# # # # #         df_f.to_csv(index=False).encode("utf-8"),
# # # # #         file_name="summary_cleaned_filtered.csv",
# # # # #         mime="text/csv",
# # # # #     )

# # # # # else:
# # # # #     st.info("Upload a file to begin.")

# # # # import streamlit as st
# # # # import pandas as pd
# # # # import numpy as np
# # # # import plotly.express as px
# # # # import plotly.graph_objects as go
# # # # import re
# # # # from scipy import stats
# # # # from datetime import date

# # # # st.set_page_config(page_title="GHG Summary Analyzer", layout="wide")
# # # # st.title("GHG Summary Analyzer")
# # # # st.caption("Upload an Excel; we'll detect the Summary sheet, clean it, and build interactive charts.")

# # # # uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

# # # # # ---------------- Helpers ----------------
# # # # def load_excel(uploaded_file):
# # # #     for eng in ("openpyxl", "calamine"):
# # # #         try:
# # # #             return pd.ExcelFile(uploaded_file, engine=eng), eng
# # # #         except Exception:
# # # #             continue
# # # #     raise ImportError("No Excel engine available.")

# # # # def find_sheet(xls):
# # # #     if "Summary Sheet" in xls.sheet_names:
# # # #         return "Summary Sheet"
# # # #     for s in xls.sheet_names:
# # # #         if "summary" in s.lower():
# # # #             return s
# # # #     return xls.sheet_names[0]

# # # # def header_detect_clean(df_raw):
# # # #     raw = df_raw.dropna(how="all").dropna(axis=1, how="all").copy()
# # # #     header_idx = raw.head(min(15, len(raw))).notna().sum(axis=1).idxmax()
# # # #     df = raw.copy()
# # # #     df.columns = df.loc[header_idx].astype(str).str.strip()
# # # #     df = df.loc[header_idx + 1:].reset_index(drop=True)
# # # #     seen, cols = {}, []
# # # #     for c in df.columns:
# # # #         base = c if c and c != "nan" else "Unnamed"
# # # #         seen[base] = seen.get(base, 0) + 1
# # # #         cols.append(base if seen[base] == 1 else f"{base}_{seen[base]}")
# # # #     df.columns = cols
# # # #     for c in df.columns:
# # # #         parsed = pd.to_numeric(df[c], errors="coerce")
# # # #         if parsed.notna().mean() >= 0.6:
# # # #             df[c] = parsed
# # # #     return df

# # # # def prune_columns(df, drop_unnamed=True, null_threshold=0.90, drop_pattern=None):
# # # #     drop = []
# # # #     for c in df.columns:
# # # #         name = str(c).strip()
# # # #         if drop_unnamed and (name.lower().startswith("unnamed") or name == "" or name.lower() == "nan"):
# # # #             drop.append(c); continue
# # # #         if drop_pattern and re.match(drop_pattern, name):
# # # #             drop.append(c); continue
# # # #         if df[c].isna().mean() >= null_threshold:
# # # #             drop.append(c); continue
# # # #         if df[c].dtype == "object":
# # # #             s = df[c].astype(str).str.strip().replace("nan", "")
# # # #             if s.replace("", np.nan).dropna().empty:
# # # #                 drop.append(c); continue
# # # #     return df.drop(columns=drop), drop

# # # # def guess_label_cols(df, num_cols):
# # # #     return [c for c in df.columns if c not in num_cols]

# # # # def arrow_safe(df):
# # # #     out = df.copy()
# # # #     for c in out.columns:
# # # #         if out[c].dtype == "object":
# # # #             out[c] = out[c].astype("string")
# # # #     return out

# # # # def hex_to_rgba(hex_color, alpha):
# # # #     try:
# # # #         hc = str(hex_color).lstrip("#")
# # # #         if len(hc) != 6:
# # # #             return hex_color
# # # #         r, g, b = int(hc[0:2], 16), int(hc[2:4], 16), int(hc[4:6], 16)
# # # #         return f"rgba({r},{g},{b},{alpha})"
# # # #     except Exception:
# # # #         return hex_color

# # # # PALETTES = {
# # # #     "Plotly": px.colors.qualitative.Plotly,
# # # #     "D3": px.colors.qualitative.D3,
# # # #     "Bold": px.colors.qualitative.Bold,
# # # #     "Dark24": px.colors.qualitative.Dark24,
# # # # }

# # # # # ─────────────────────────────────────────────────────────────────
# # # # # GHG PARSER
# # # # # ─────────────────────────────────────────────────────────────────
# # # # GHG_ROW_LABELS = {
# # # #     "Emission per ton of Equivalent product":           "T",
# # # #     "Total Direct and Indirect Emission":               "S",
# # # #     "Total Direct Emission (Scope 1)":                  "Scope1",
# # # #     "Total Indirect Emission (Scope 2)":                "Scope2",
# # # #     "Total Equivalent Product for GHG Emission Intensity": "P",
# # # # }

# # # # KNOWN_MONTHS = ["January","February","March","April","May","June",
# # # #                 "July","August","September","October","November","December"]

# # # # # Indian fiscal year: April=start, March=end
# # # # FISCAL_YEAR_ORDER = ["April","May","June","July","August","September",
# # # #                      "October","November","December","January","February","March"]


# # # # def get_remaining_fiscal_months(months_known):
# # # #     """Return months in the fiscal year that are AFTER the last known month."""
# # # #     if not months_known:
# # # #         return []
# # # #     last_known = months_known[-1]
# # # #     if last_known not in FISCAL_YEAR_ORDER:
# # # #         return []
# # # #     last_idx = FISCAL_YEAR_ORDER.index(last_known)
# # # #     return FISCAL_YEAR_ORDER[last_idx + 1:]


# # # # def extract_ghg_data(raw):
# # # #     month_cols = {}
# # # #     for i, row in raw.iterrows():
# # # #         for col_idx, val in enumerate(row):
# # # #             vs = str(val).strip()
# # # #             for m in KNOWN_MONTHS:
# # # #                 if m.lower() == vs.lower():
# # # #                     month_cols[m] = col_idx
# # # #         if len(month_cols) >= 2:
# # # #             break

# # # #     results = {label: {} for label in GHG_ROW_LABELS}
# # # #     for i, row in raw.iterrows():
# # # #         row_vals = [str(v).strip() for v in row]
# # # #         for label in GHG_ROW_LABELS:
# # # #             if any(label.lower() in rv.lower() for rv in row_vals):
# # # #                 for month, col_idx in month_cols.items():
# # # #                     try:
# # # #                         v = pd.to_numeric(row.iloc[col_idx], errors="coerce")
# # # #                         if pd.notna(v):
# # # #                             results[label][month] = v
# # # #                     except Exception:
# # # #                         pass

# # # #     # Extract baseline from last row (T row) last non-null column if present
# # # #     ordered_months = [m for m in FISCAL_YEAR_ORDER if m in month_cols]
# # # #     return results, ordered_months


# # # # def extract_baseline_from_raw(raw):
# # # #     """Try to get baseline value from the T row's 'Current year' or last column."""
# # # #     for i, row in raw.iterrows():
# # # #         row_vals = [str(v).strip() for v in row]
# # # #         if any("emission per ton" in rv.lower() for rv in row_vals):
# # # #             # last non-nan numeric value in that row
# # # #             numeric_vals = pd.to_numeric(pd.Series(row.tolist()), errors="coerce").dropna()
# # # #             if not numeric_vals.empty:
# # # #                 return float(numeric_vals.iloc[-1])
# # # #     return None


# # # # def predict_remaining_months(months_known, values, all_months):
# # # #     if len(values) < 2:
# # # #         return {}, 0, 0
# # # #     x_known = np.array([all_months.index(m) for m in months_known if m in all_months])
# # # #     y_known = np.array(values)
# # # #     slope, intercept, *_ = stats.linregress(x_known, y_known)
# # # #     predictions = {}
# # # #     for m in all_months:
# # # #         if m not in months_known:
# # # #             predictions[m] = slope * all_months.index(m) + intercept
# # # #     return predictions, slope, intercept


# # # # # ─────────────────────────────────────────────────────────────────
# # # # # DEDICATED GEI CHART (Emission per ton) — green baseline line
# # # # # ─────────────────────────────────────────────────────────────────
# # # # def make_gei_chart(ghg_data, months_known, baseline_value, chart_type, add_prediction):
# # # #     label = "Emission per ton of Equivalent product"
# # # #     color = "#1976D2"
# # # #     all_months = FISCAL_YEAR_ORDER

# # # #     actual_y = [ghg_data[label].get(m, None) for m in all_months]

# # # #     pred_dict = {}
# # # #     if add_prediction and len(months_known) >= 2:
# # # #         known_vals = [ghg_data[label][m] for m in months_known if m in ghg_data[label]]
# # # #         pred_result = predict_remaining_months(months_known, known_vals, all_months)
# # # #         if isinstance(pred_result, tuple):
# # # #             pred_dict = pred_result[0]

# # # #     fig = go.Figure()
# # # #     rgba_fill = hex_to_rgba(color, 0.22)

# # # #     if chart_type == "Bar":
# # # #         fig.add_trace(go.Bar(
# # # #             x=all_months, y=[v if v is not None else 0 for v in actual_y],
# # # #             name="Actual (GEI)",
# # # #             marker_color=[color if v is not None else "rgba(0,0,0,0)" for v in actual_y],
# # # #             text=[f"{v:.4f}" if v is not None else "" for v in actual_y],
# # # #             textposition="outside",
# # # #         ))
# # # #     elif chart_type == "Line":
# # # #         fig.add_trace(go.Scatter(
# # # #             x=all_months, y=actual_y, name="Actual (GEI)",
# # # #             mode="lines+markers+text",
# # # #             line=dict(color=color, width=2.5),
# # # #             marker=dict(size=8),
# # # #             text=[f"{v:.4f}" if v is not None else "" for v in actual_y],
# # # #             textposition="top center",
# # # #         ))
# # # #     else:
# # # #         fig.add_trace(go.Scatter(
# # # #             x=all_months, y=actual_y, name="Actual (GEI)",
# # # #             mode="lines+markers", fill="tozeroy",
# # # #             line=dict(color=color, width=2.5), fillcolor=rgba_fill,
# # # #             marker=dict(size=8),
# # # #         ))

# # # #     if pred_dict:
# # # #         pred_x = [m for m in all_months if m in pred_dict]
# # # #         pred_vals = [pred_dict[m] for m in pred_x]
# # # #         fig.add_trace(go.Scatter(
# # # #             x=pred_x, y=pred_vals, name="Predicted (Trend)",
# # # #             mode="lines+markers+text",
# # # #             line=dict(color="#FF9800", width=2, dash="dash"),
# # # #             marker=dict(size=8, symbol="diamond"),
# # # #             text=[f"{v:.4f}" for v in pred_vals],
# # # #             textposition="top center",
# # # #         ))

# # # #     # ── GREEN baseline line ─────────────────────────────────────
# # # #     fig.add_hline(
# # # #         y=baseline_value,
# # # #         line=dict(color="green", width=2.5, dash="solid"),
# # # #         annotation_text=f"Baseline Target: {baseline_value:.4f}",
# # # #         annotation_position="top right",
# # # #         annotation_font=dict(color="green", size=13, family="Arial Black"),
# # # #     )

# # # #     fig.update_layout(
# # # #         title=dict(
# # # #             text="📌 GHG Emission Intensity — Emission per ton of Equivalent Product",
# # # #             font=dict(size=16)
# # # #         ),
# # # #         xaxis_title="Month (Fiscal Year)",
# # # #         yaxis_title="tCO₂ eq / ton of equivalent product",
# # # #         height=500,
# # # #         margin=dict(l=10, r=10, t=60, b=10),
# # # #         hovermode="x unified",
# # # #         legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
# # # #         plot_bgcolor="rgba(240,248,255,0.4)",
# # # #     )
# # # #     return fig


# # # # def make_ghg_chart(ghg_data, label, months_known, chart_type, y_axis_label, color):
# # # #     all_months = FISCAL_YEAR_ORDER
# # # #     actual_y = [ghg_data[label].get(m, None) for m in all_months]
# # # #     rgba_fill = hex_to_rgba(color, 0.22)

# # # #     fig = go.Figure()
# # # #     if chart_type == "Bar":
# # # #         fig.add_trace(go.Bar(
# # # #             x=all_months, y=[v if v is not None else 0 for v in actual_y],
# # # #             name="Actual",
# # # #             marker_color=[color if v is not None else "rgba(0,0,0,0)" for v in actual_y],
# # # #             text=[f"{v:,.2f}" if v is not None else "" for v in actual_y],
# # # #             textposition="outside",
# # # #         ))
# # # #     elif chart_type == "Line":
# # # #         fig.add_trace(go.Scatter(
# # # #             x=all_months, y=actual_y, name="Actual",
# # # #             mode="lines+markers+text",
# # # #             line=dict(color=color, width=2.5), marker=dict(size=8),
# # # #             text=[f"{v:,.2f}" if v is not None else "" for v in actual_y],
# # # #             textposition="top center",
# # # #         ))
# # # #     else:
# # # #         fig.add_trace(go.Scatter(
# # # #             x=all_months, y=actual_y, name="Actual",
# # # #             mode="lines+markers", fill="tozeroy",
# # # #             line=dict(color=color, width=2.5), fillcolor=rgba_fill,
# # # #             marker=dict(size=8),
# # # #         ))

# # # #     fig.update_layout(
# # # #         title=dict(text=label, font=dict(size=15)),
# # # #         xaxis_title="Month (Fiscal Year)",
# # # #         yaxis_title=y_axis_label,
# # # #         height=460,
# # # #         margin=dict(l=10, r=10, t=50, b=10),
# # # #         hovermode="x unified",
# # # #         legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
# # # #     )
# # # #     return fig


# # # # # ─────────────────────────────────────────────────────────────────
# # # # # APP
# # # # # ─────────────────────────────────────────────────────────────────
# # # # if uploaded:
# # # #     try:
# # # #         xls, engine_used = load_excel(uploaded)
# # # #         st.caption(f"Excel engine: **{engine_used}**")
# # # #         sheet = find_sheet(xls)
# # # #         st.success(f"Using sheet: **{sheet}**")
# # # #         raw = pd.read_excel(uploaded, sheet_name=sheet, engine=engine_used, header=None)
# # # #         df = header_detect_clean(raw)
# # # #     except Exception as e:
# # # #         st.exception(e); st.stop()

# # # #     drop_pattern = r"^0\.0_.*"
# # # #     df, dropped_prune = prune_columns(df, drop_unnamed=True, null_threshold=0.90, drop_pattern=drop_pattern)
# # # #     if dropped_prune:
# # # #         st.info(f"Hidden {len(dropped_prune)} mostly-empty / unnamed / patterned columns.")

# # # #     num_cols = df.select_dtypes(include="number").columns.tolist()
# # # #     valid_y_cols, excluded_y_cols = [], []
# # # #     for c in num_cols:
# # # #         s = pd.to_numeric(df[c], errors="coerce").dropna()
# # # #         (valid_y_cols if (not s.empty and s.abs().sum() != 0) else excluded_y_cols).append(c)

# # # #     cat_cols = guess_label_cols(df, num_cols)
# # # #     ghg_data, months_known = extract_ghg_data(raw)

# # # #     # Try to auto-detect baseline from xlsx; fallback to 0.3552
# # # #     auto_baseline = extract_baseline_from_raw(raw)

# # # #     # ─────────────────────────────────────────────────────────────
# # # #     # SIDEBAR
# # # #     # ─────────────────────────────────────────────────────────────
# # # #     with st.sidebar:
# # # #         st.header("Controls")
# # # #         x_col = st.selectbox("X-axis (categorical / time)", options=(cat_cols or [None]))
# # # #         if valid_y_cols:
# # # #             y_cols = st.multiselect("Y-axis (numeric)", options=valid_y_cols, default=valid_y_cols[:1])
# # # #         else:
# # # #             st.warning("No numeric columns with non-zero/nonnull values detected.")
# # # #             y_cols = []

# # # #         if y_cols:
# # # #             y_numeric_all = df[y_cols].apply(pd.to_numeric, errors="coerce").fillna(0).abs()
# # # #             df_for_filters = df.loc[(y_numeric_all.sum(axis=1) != 0)].copy()
# # # #         else:
# # # #             df_for_filters = df.copy()

# # # #         st.markdown("**Filters**")
# # # #         active_filters = {}
# # # #         for c in cat_cols[:6]:
# # # #             vals = df_for_filters[c].dropna().astype(str).unique().tolist()
# # # #             if 1 < len(vals) <= 200:
# # # #                 sel = st.multiselect(f"{c}", options=vals, default=[], placeholder="(all)")
# # # #                 if sel:
# # # #                     active_filters[c] = set(sel)

# # # #         chart_type = st.radio("Chart type (generic)", ["Line", "Bar", "Area"], index=1)

# # # #         st.markdown("---")
# # # #         st.subheader("GHG Chart Options")
# # # #         ghg_chart_type = st.radio("GHG Chart type", ["Line", "Bar", "Area"], index=0)
# # # #         show_prediction = st.checkbox("Show Predicted Months", value=True)

# # # #         default_bl = auto_baseline if auto_baseline else 0.3552
# # # #         baseline_value = st.number_input(
# # # #             "Baseline Target (green line on GEI chart)",
# # # #             value=default_bl, step=0.001, format="%.4f",
# # # #             help="Solid green reference line on Emission Intensity chart"
# # # #         )

# # # #     # ─────────────────────────────────────────────────────────────
# # # #     # FILTERS
# # # #     # ─────────────────────────────────────────────────────────────
# # # #     filt = pd.Series(True, index=df.index)
# # # #     for col, allowed in active_filters.items():
# # # #         filt &= df[col].astype(str).isin(allowed)
# # # #     df_f = df.loc[filt].copy()

# # # #     if y_cols:
# # # #         y_numeric_after = df_f[y_cols].apply(pd.to_numeric, errors="coerce")
# # # #         nonzero_mask = (y_numeric_after.fillna(0).abs().sum(axis=1) != 0)
# # # #         removed = (~nonzero_mask).sum()
# # # #         if removed:
# # # #             st.info(f"Removed {int(removed)} row(s) where selected Y columns were all zero/null.")
# # # #         df_f = df_f.loc[nonzero_mask].reset_index(drop=True)

# # # #     if excluded_y_cols:
# # # #         st.info(f"Excluded numeric columns (all zero or all null): {', '.join(excluded_y_cols)}")
# # # #     if df_f.empty:
# # # #         st.warning("No rows left after filters. Adjust selection."); st.stop()

# # # #     st.subheader("Data Table")
# # # #     st.dataframe(arrow_safe(df_f), use_container_width=True, height=420)

# # # #     # ─────────────────────────────────────────────────────────────
# # # #     # ★ GHG DEDICATED SECTION ★
# # # #     # ─────────────────────────────────────────────────────────────
# # # #     any_ghg = any(bool(ghg_data[lbl]) for lbl in GHG_ROW_LABELS)

# # # #     if any_ghg and months_known:
# # # #         st.markdown("---")
# # # #         st.header("📊 GHG Emission Analysis")

# # # #         # Metric cards for known months (GEI row)
# # # #         gei_label = "Emission per ton of Equivalent product"
# # # #         col_info = st.columns(len(months_known))
# # # #         for i, m in enumerate(months_known):
# # # #             v = ghg_data[gei_label].get(m)
# # # #             delta_val = round(v - baseline_value, 4) if v is not None else None
# # # #             col_info[i].metric(
# # # #                 label=m,
# # # #                 value=f"{v:.4f}" if v is not None else "N/A",
# # # #                 delta=f"{delta_val:+.4f} vs baseline" if delta_val is not None else None,
# # # #                 delta_color="inverse",
# # # #             )

# # # #         # ── 1. DEDICATED GEI CHART (solo, prominent) ─────────────
# # # #         st.markdown("---")
# # # #         st.subheader("🎯 Emission Intensity per Ton — with Green Baseline Target")
# # # #         if ghg_data.get(gei_label):
# # # #             fig_gei = make_gei_chart(
# # # #                 ghg_data=ghg_data,
# # # #                 months_known=months_known,
# # # #                 baseline_value=baseline_value,
# # # #                 chart_type=ghg_chart_type,
# # # #                 add_prediction=show_prediction,
# # # #             )
# # # #             st.plotly_chart(fig_gei, use_container_width=True)
# # # #         else:
# # # #             st.warning("No data for Emission per ton of Equivalent product.")

# # # #         # ── 2. OTHER GHG CHARTS (each separate, different colors) ─
# # # #         OTHER_CHART_CONFIGS = [
# # # #             {
# # # #                 "label": "Total Direct and Indirect Emission",
# # # #                 "y_axis": "tCO₂ eq",
# # # #                 "color": "#E53935",
# # # #             },
# # # #             {
# # # #                 "label": "Total Direct Emission (Scope 1)",
# # # #                 "y_axis": "tCO₂ eq",
# # # #                 "color": "#FF7043",
# # # #             },
# # # #             {
# # # #                 "label": "Total Indirect Emission (Scope 2)",
# # # #                 "y_axis": "tCO₂ eq",
# # # #                 "color": "#7B1FA2",
# # # #             },
# # # #             {
# # # #                 "label": "Total Equivalent Product for GHG Emission Intensity",
# # # #                 "y_axis": "Tonnes",
# # # #                 "color": "#2E7D32",
# # # #             },
# # # #         ]

# # # #         st.markdown("---")
# # # #         st.subheader("📈 Other GHG Metrics")
# # # #         for cfg in OTHER_CHART_CONFIGS:
# # # #             lbl = cfg["label"]
# # # #             if not ghg_data.get(lbl):
# # # #                 st.warning(f"No data found for: {lbl}")
# # # #                 continue
# # # #             fig = make_ghg_chart(
# # # #                 ghg_data=ghg_data,
# # # #                 label=lbl,
# # # #                 months_known=months_known,
# # # #                 chart_type=ghg_chart_type,
# # # #                 y_axis_label=cfg["y_axis"],
# # # #                 color=cfg["color"],
# # # #             )
# # # #             st.subheader(lbl)
# # # #             st.plotly_chart(fig, use_container_width=True)

# # # #         # ── 3. PREDICTION TABLE — remaining fiscal months only ─────
# # # #         if show_prediction:
# # # #             remaining = get_remaining_fiscal_months(months_known)

# # # #             st.markdown("---")
# # # #             st.subheader("📅 Baseline Prediction — Remaining Months of This Fiscal Year")

# # # #             if not remaining:
# # # #                 st.info("No remaining months to predict — all fiscal year months have data.")
# # # #             else:
# # # #                 st.caption(
# # # #                     f"Months with data: **{', '.join(months_known)}**  |  "
# # # #                     f"Predicting: **{', '.join(remaining)}**"
# # # #                 )
# # # #                 pred_rows = []
# # # #                 for lbl in GHG_ROW_LABELS:
# # # #                     vals = ghg_data[lbl]
# # # #                     if len(vals) < 2:
# # # #                         continue
# # # #                     mk = [m for m in FISCAL_YEAR_ORDER if m in vals]
# # # #                     mv = [vals[m] for m in mk]
# # # #                     try:
# # # #                         pred_dict, slope, intercept = predict_remaining_months(mk, mv, FISCAL_YEAR_ORDER)
# # # #                         for m in remaining:          # ← only remaining fiscal months
# # # #                             if m in pred_dict:
# # # #                                 pred_rows.append({
# # # #                                     "Metric": lbl,
# # # #                                     "Month": m,
# # # #                                     "Predicted Value": round(pred_dict[m], 6),
# # # #                                 })
# # # #                     except Exception:
# # # #                         pass

# # # #                 if pred_rows:
# # # #                     pred_df = pd.DataFrame(pred_rows)
# # # #                     # Pivot for readability
# # # #                     pivot = pred_df.pivot(index="Metric", columns="Month", values="Predicted Value")
# # # #                     # Reorder columns to fiscal year order
# # # #                     pivot = pivot[[m for m in remaining if m in pivot.columns]]
# # # #                     st.dataframe(pivot, use_container_width=True)
# # # #                 else:
# # # #                     st.info("Not enough data to generate predictions.")

# # # #     # ─────────────────────────────────────────────────────────────
# # # #     # GENERIC CHARTS
# # # #     # ─────────────────────────────────────────────────────────────
# # # #     if not y_cols:
# # # #         st.warning("Pick at least one numeric Y column."); st.stop()

# # # #     st.markdown("---")
# # # #     st.header("📈 Generic Charts (Selected Columns)")

# # # #     if x_col and x_col in df_f.columns:
# # # #         x_vals = df_f[x_col].astype(str).fillna("")
# # # #     else:
# # # #         x_vals = pd.Series(range(len(df_f)), index=df_f.index, name="Index")
# # # #         x_col = "Index"
# # # #         df_f = df_f.assign(Index=x_vals)

# # # #     for i, y in enumerate(y_cols):
# # # #         primary_color = px.colors.qualitative.Plotly[i % len(px.colors.qualitative.Plotly)]
# # # #         y_vals = pd.to_numeric(df_f[y], errors="coerce")
# # # #         if y_vals.dropna().empty or y_vals.dropna().abs().sum() == 0:
# # # #             st.info(f"Skipping '{y}' (all zero/null)."); continue
# # # #         df_plot = pd.DataFrame({x_col: x_vals, y: y_vals.fillna(0)})

# # # #         if chart_type == "Bar":
# # # #             fig = px.bar(df_plot, x=x_col, y=y)
# # # #             fig.update_traces(marker=dict(color=primary_color))
# # # #         elif chart_type == "Line":
# # # #             fig = px.line(df_plot, x=x_col, y=y)
# # # #             fig.update_traces(line=dict(color=primary_color))
# # # #         else:
# # # #             fig = px.area(df_plot, x=x_col, y=y)
# # # #             fig.update_traces(line=dict(color=primary_color),
# # # #                               fillcolor=hex_to_rgba(primary_color, 0.25))

# # # #         fig.update_layout(margin=dict(l=10, r=10, t=40, b=10), height=480, hovermode="x unified")
# # # #         st.plotly_chart(fig, use_container_width=True)

# # # #     st.download_button(
# # # #         "Download cleaned/filtered data (CSV)",
# # # #         df_f.to_csv(index=False).encode("utf-8"),
# # # #         file_name="summary_cleaned_filtered.csv",
# # # #         mime="text/csv",
# # # #     )

# # # # else:
# # # #     st.info("Upload a file to begin.")

# # # import streamlit as st
# # # import pandas as pd
# # # import numpy as np
# # # import plotly.express as px
# # # import plotly.graph_objects as go
# # # import re
# # # from scipy import stats
# # # from datetime import date

# # # st.set_page_config(page_title="GHG Summary Analyzer", layout="wide")
# # # st.title("GHG Summary Analyzer")
# # # st.caption("Upload an Excel; we'll detect the Summary sheet, clean it, and build interactive charts.")

# # # uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

# # # # ---------------- Helpers ----------------
# # # def load_excel(uploaded_file):
# # #     for eng in ("openpyxl", "calamine"):
# # #         try:
# # #             return pd.ExcelFile(uploaded_file, engine=eng), eng
# # #         except Exception:
# # #             continue
# # #     raise ImportError("No Excel engine available.")

# # # def find_sheet(xls):
# # #     if "Summary Sheet" in xls.sheet_names:
# # #         return "Summary Sheet"
# # #     for s in xls.sheet_names:
# # #         if "summary" in s.lower():
# # #             return s
# # #     return xls.sheet_names[0]

# # # def header_detect_clean(df_raw):
# # #     raw = df_raw.dropna(how="all").dropna(axis=1, how="all").copy()
# # #     header_idx = raw.head(min(15, len(raw))).notna().sum(axis=1).idxmax()
# # #     df = raw.copy()
# # #     df.columns = df.loc[header_idx].astype(str).str.strip()
# # #     df = df.loc[header_idx + 1:].reset_index(drop=True)
# # #     seen, cols = {}, []
# # #     for c in df.columns:
# # #         base = c if c and c != "nan" else "Unnamed"
# # #         seen[base] = seen.get(base, 0) + 1
# # #         cols.append(base if seen[base] == 1 else f"{base}_{seen[base]}")
# # #     df.columns = cols
# # #     for c in df.columns:
# # #         parsed = pd.to_numeric(df[c], errors="coerce")
# # #         if parsed.notna().mean() >= 0.6:
# # #             df[c] = parsed
# # #     return df

# # # def prune_columns(df, drop_unnamed=True, null_threshold=0.90, drop_pattern=None):
# # #     drop = []
# # #     for c in df.columns:
# # #         name = str(c).strip()
# # #         if drop_unnamed and (name.lower().startswith("unnamed") or name == "" or name.lower() == "nan"):
# # #             drop.append(c); continue
# # #         if drop_pattern and re.match(drop_pattern, name):
# # #             drop.append(c); continue
# # #         if df[c].isna().mean() >= null_threshold:
# # #             drop.append(c); continue
# # #         if df[c].dtype == "object":
# # #             s = df[c].astype(str).str.strip().replace("nan", "")
# # #             if s.replace("", np.nan).dropna().empty:
# # #                 drop.append(c); continue
# # #     return df.drop(columns=drop), drop

# # # def guess_label_cols(df, num_cols):
# # #     return [c for c in df.columns if c not in num_cols]

# # # def arrow_safe(df):
# # #     out = df.copy()
# # #     for c in out.columns:
# # #         if out[c].dtype == "object":
# # #             out[c] = out[c].astype("string")
# # #     return out

# # # def hex_to_rgba(hex_color, alpha):
# # #     try:
# # #         hc = str(hex_color).lstrip("#")
# # #         if len(hc) != 6:
# # #             return hex_color
# # #         r, g, b = int(hc[0:2], 16), int(hc[2:4], 16), int(hc[4:6], 16)
# # #         return f"rgba({r},{g},{b},{alpha})"
# # #     except Exception:
# # #         return hex_color

# # # PALETTES = {
# # #     "Plotly": px.colors.qualitative.Plotly,
# # #     "D3": px.colors.qualitative.D3,
# # #     "Bold": px.colors.qualitative.Bold,
# # #     "Dark24": px.colors.qualitative.Dark24,
# # # }

# # # # ─────────────────────────────────────────────────────────────────
# # # # GHG PARSER
# # # # ─────────────────────────────────────────────────────────────────
# # # GHG_ROW_LABELS = {
# # #     "Emission per ton of Equivalent product":           "T",
# # #     "Total Direct and Indirect Emission":               "S",
# # #     "Total Direct Emission (Scope 1)":                  "Scope1",
# # #     "Total Indirect Emission (Scope 2)":                "Scope2",
# # #     "Total Equivalent Product for GHG Emission Intensity": "P",
# # # }

# # # KNOWN_MONTHS = ["January","February","March","April","May","June",
# # #                 "July","August","September","October","November","December"]

# # # # Indian fiscal year: April=start, March=end
# # # FISCAL_YEAR_ORDER = ["April","May","June","July","August","September",
# # #                      "October","November","December","January","February","March"]


# # # def get_remaining_fiscal_months(months_known):
# # #     """Return months in the fiscal year that are AFTER the last known month."""
# # #     if not months_known:
# # #         return []
# # #     last_known = months_known[-1]
# # #     if last_known not in FISCAL_YEAR_ORDER:
# # #         return []
# # #     last_idx = FISCAL_YEAR_ORDER.index(last_known)
# # #     return FISCAL_YEAR_ORDER[last_idx + 1:]


# # # def extract_ghg_data(raw):
# # #     month_cols = {}
# # #     for i, row in raw.iterrows():
# # #         for col_idx, val in enumerate(row):
# # #             vs = str(val).strip()
# # #             for m in KNOWN_MONTHS:
# # #                 if m.lower() == vs.lower():
# # #                     month_cols[m] = col_idx
# # #         if len(month_cols) >= 2:
# # #             break

# # #     results = {label: {} for label in GHG_ROW_LABELS}
# # #     for i, row in raw.iterrows():
# # #         row_vals = [str(v).strip() for v in row]
# # #         for label in GHG_ROW_LABELS:
# # #             if any(label.lower() in rv.lower() for rv in row_vals):
# # #                 for month, col_idx in month_cols.items():
# # #                     try:
# # #                         v = pd.to_numeric(row.iloc[col_idx], errors="coerce")
# # #                         if pd.notna(v):
# # #                             results[label][month] = v
# # #                     except Exception:
# # #                         pass

# # #     ordered_months = [m for m in FISCAL_YEAR_ORDER if m in month_cols]
# # #     return results, ordered_months


# # # def extract_baseline_from_raw(raw):
# # #     """Try to get baseline value from the T row's last non-null column if present."""
# # #     for i, row in raw.iterrows():
# # #         row_vals = [str(v).strip() for v in row]
# # #         if any("emission per ton" in rv.lower() for rv in row_vals):
# # #             numeric_vals = pd.to_numeric(pd.Series(row.tolist()), errors="coerce").dropna()
# # #             if not numeric_vals.empty:
# # #                 return float(numeric_vals.iloc[-1])
# # #     return None


# # # def predict_remaining_months(months_known, values, all_months):
# # #     if len(values) < 2:
# # #         return {}, 0, 0
# # #     x_known = np.array([all_months.index(m) for m in months_known if m in all_months])
# # #     y_known = np.array(values)
# # #     slope, intercept, *_ = stats.linregress(x_known, y_known)
# # #     predictions = {}
# # #     for m in all_months:
# # #         if m not in months_known:
# # #             predictions[m] = slope * all_months.index(m) + intercept
# # #     return predictions, slope, intercept


# # # # ─────────────────────────────────────────────────────────────────
# # # # DEDICATED GEI CHART
# # # # Two straight reference lines: T (green) and Target (red)
# # # # Wave fill between the two lines
# # # # Actual GEI as a smooth green wave line
# # # # ─────────────────────────────────────────────────────────────────
# # # def make_gei_chart(ghg_data, months_known, baseline_value, t_line_value, chart_type, add_prediction):
# # #     label = "Emission per ton of Equivalent product"
# # #     all_months = FISCAL_YEAR_ORDER

# # #     # ── Build actual + predicted green wave ──────────────────────
# # #     pred_dict = {}
# # #     if len(months_known) >= 2:
# # #         known_vals = [ghg_data[label][m] for m in months_known if m in ghg_data[label]]
# # #         pred_result = predict_remaining_months(months_known, known_vals, all_months)
# # #         if isinstance(pred_result, tuple):
# # #             pred_dict = pred_result[0]

# # #     green_y = []
# # #     green_mode = []
# # #     for m in all_months:
# # #         if m in ghg_data[label] and ghg_data[label][m] is not None:
# # #             green_y.append(ghg_data[label][m])
# # #             green_mode.append("actual")
# # #         elif add_prediction and m in pred_dict:
# # #             green_y.append(pred_dict[m])
# # #             green_mode.append("predicted")
# # #         else:
# # #             green_y.append(None)
# # #             green_mode.append("none")

# # #     # Flat reference lines across all 12 months
# # #     target_y = [baseline_value] * len(all_months)   # RED line (Target)
# # #     t_line_y = [t_line_value] * len(all_months)     # GREEN straight line (T)

# # #     fig = go.Figure()

# # #     # ── BAND FILL between T line and Target line ─────────────────
# # #     # We fill "between" the two straight lines using tonexty
# # #     # Lower boundary first (T line), then upper boundary (Target)
# # #     # Determine which is higher dynamically
# # #     lower_val = min(t_line_value, baseline_value)
# # #     upper_val = max(t_line_value, baseline_value)
# # #     lower_y = [lower_val] * len(all_months)
# # #     upper_y = [upper_val] * len(all_months)

# # #     # Lower bound trace (invisible line, used as fill base)
# # #     fig.add_trace(go.Scatter(
# # #         x=all_months,
# # #         y=lower_y,
# # #         mode="lines",
# # #         line=dict(color="rgba(0,0,0,0)", width=0),
# # #         showlegend=False,
# # #         hoverinfo="skip",
# # #         name="_lower_bound",
# # #     ))

# # #     # Upper bound trace — fill tonexty creates the band
# # #     fig.add_trace(go.Scatter(
# # #         x=all_months,
# # #         y=upper_y,
# # #         mode="lines",
# # #         line=dict(color="rgba(0,0,0,0)", width=0),
# # #         fill="tonexty",
# # #         fillcolor="rgba(100,220,130,0.18)",
# # #         showlegend=True,
# # #         name="Zone between T & Target",
# # #         hoverinfo="skip",
# # #     ))

# # #     # ── ACTUAL GEI wave line (solid green) ───────────────────────
# # #     actual_x = [m for m, mode in zip(all_months, green_mode) if mode == "actual"]
# # #     actual_vals = [v for v, mode in zip(green_y, green_mode) if mode == "actual"]

# # #     if actual_x:
# # #         fig.add_trace(go.Scatter(
# # #             x=actual_x,
# # #             y=actual_vals,
# # #             name="Actual GEI",
# # #             mode="lines+markers+text",
# # #             line=dict(color="#00E676", width=3, shape="spline", smoothing=1.2),
# # #             marker=dict(size=9, color="#00E676", line=dict(color="white", width=1.5)),
# # #             text=[f"{v:.4f}" for v in actual_vals],
# # #             textposition="top center",
# # #             textfont=dict(color="#00E676", size=11),
# # #         ))

# # #     # ── PREDICTED GEI wave line (dashed green) ───────────────────
# # #     pred_x = [m for m, mode in zip(all_months, green_mode) if mode == "predicted"]
# # #     pred_vals = [v for v, mode in zip(green_y, green_mode) if mode == "predicted"]
# # #     if pred_x:
# # #         bridge_x = actual_x[-1:] + pred_x[:1] if actual_x else pred_x[:1]
# # #         bridge_y = actual_vals[-1:] + pred_vals[:1] if actual_vals else pred_vals[:1]
# # #         fig.add_trace(go.Scatter(
# # #             x=bridge_x, y=bridge_y,
# # #             mode="lines", showlegend=False,
# # #             line=dict(color="#00E676", width=2, dash="dot", shape="spline"),
# # #             hoverinfo="skip",
# # #         ))
# # #         fig.add_trace(go.Scatter(
# # #             x=pred_x, y=pred_vals,
# # #             name="Predicted GEI (Trend)",
# # #             mode="lines+markers+text",
# # #             line=dict(color="#00E676", width=2.5, dash="dot", shape="spline", smoothing=1.2),
# # #             marker=dict(size=8, color="#00E676", symbol="diamond",
# # #                         line=dict(color="white", width=1)),
# # #             text=[f"{v:.4f}" for v in pred_vals],
# # #             textposition="top center",
# # #             textfont=dict(color="#00E676", size=10),
# # #         ))

# # #     # ── T LINE — green straight line ─────────────────────────────
# # #     fig.add_trace(go.Scatter(
# # #         x=all_months,
# # #         y=t_line_y,
# # #         name=f"Emission per ton of Equivalent product (T) = {t_line_value:.6f}",
# # #         mode="lines+text",
# # #         line=dict(color="#43A047", width=2.5, shape="linear", dash="solid"),
# # #         text=["" if i < len(all_months) - 1 else f"T = {t_line_value:.4f}" for i in range(len(all_months))],
# # #         textposition="middle right",
# # #         textfont=dict(color="#43A047", size=12, family="monospace"),
# # #     ))

# # #     # ── TARGET LINE — red straight line ──────────────────────────
# # #     fig.add_trace(go.Scatter(
# # #         x=all_months,
# # #         y=target_y,
# # #         name=f"Baseline Target = {baseline_value:.4f}",
# # #         mode="lines+text",
# # #         line=dict(color="#F44336", width=2.5, shape="linear", dash="dash"),
# # #         text=["" if i < len(all_months) - 1 else f"Target = {baseline_value:.4f}" for i in range(len(all_months))],
# # #         textposition="middle right",
# # #         textfont=dict(color="#F44336", size=12, family="monospace"),
# # #     ))

# # #     fig.update_layout(
# # #         title=dict(
# # #             text="📌 GHG Emission Intensity — Actual GEI · T Line (Green) · Baseline Target (Red)",
# # #             font=dict(size=15, color="#EEEEEE"),
# # #         ),
# # #         xaxis=dict(
# # #             title="Month (Fiscal Year Apr → Mar)",
# # #             tickfont=dict(size=12),
# # #             gridcolor="rgba(200,200,200,0.12)",
# # #         ),
# # #         yaxis=dict(
# # #             title="tCO₂ eq / ton of equivalent product",
# # #             gridcolor="rgba(200,200,200,0.12)",
# # #             zeroline=False,
# # #         ),
# # #         height=540,
# # #         margin=dict(l=10, r=10, t=70, b=10),
# # #         hovermode="x unified",
# # #         legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
# # #                     font=dict(size=11)),
# # #         plot_bgcolor="rgba(15,20,30,0.9)",
# # #         paper_bgcolor="rgba(0,0,0,0)",
# # #     )
# # #     return fig


# # # def make_ghg_chart(ghg_data, label, months_known, chart_type, y_axis_label, color):
# # #     all_months = FISCAL_YEAR_ORDER
# # #     actual_y = [ghg_data[label].get(m, None) for m in all_months]
# # #     rgba_fill = hex_to_rgba(color, 0.22)

# # #     fig = go.Figure()
# # #     if chart_type == "Bar":
# # #         fig.add_trace(go.Bar(
# # #             x=all_months, y=[v if v is not None else 0 for v in actual_y],
# # #             name="Actual",
# # #             marker_color=[color if v is not None else "rgba(0,0,0,0)" for v in actual_y],
# # #             text=[f"{v:,.2f}" if v is not None else "" for v in actual_y],
# # #             textposition="outside",
# # #         ))
# # #     elif chart_type == "Line":
# # #         fig.add_trace(go.Scatter(
# # #             x=all_months, y=actual_y, name="Actual",
# # #             mode="lines+markers+text",
# # #             line=dict(color=color, width=2.5), marker=dict(size=8),
# # #             text=[f"{v:,.2f}" if v is not None else "" for v in actual_y],
# # #             textposition="top center",
# # #         ))
# # #     else:
# # #         fig.add_trace(go.Scatter(
# # #             x=all_months, y=actual_y, name="Actual",
# # #             mode="lines+markers", fill="tozeroy",
# # #             line=dict(color=color, width=2.5), fillcolor=rgba_fill,
# # #             marker=dict(size=8),
# # #         ))

# # #     fig.update_layout(
# # #         title=dict(text=label, font=dict(size=15)),
# # #         xaxis_title="Month (Fiscal Year)",
# # #         yaxis_title=y_axis_label,
# # #         height=460,
# # #         margin=dict(l=10, r=10, t=50, b=10),
# # #         hovermode="x unified",
# # #         legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
# # #     )
# # #     return fig


# # # # ─────────────────────────────────────────────────────────────────
# # # # APP
# # # # ─────────────────────────────────────────────────────────────────
# # # if uploaded:
# # #     try:
# # #         xls, engine_used = load_excel(uploaded)
# # #         st.caption(f"Excel engine: **{engine_used}**")
# # #         sheet = find_sheet(xls)
# # #         st.success(f"Using sheet: **{sheet}**")
# # #         raw = pd.read_excel(uploaded, sheet_name=sheet, engine=engine_used, header=None)
# # #         df = header_detect_clean(raw)
# # #     except Exception as e:
# # #         st.exception(e); st.stop()

# # #     drop_pattern = r"^0\.0_.*"
# # #     df, dropped_prune = prune_columns(df, drop_unnamed=True, null_threshold=0.90, drop_pattern=drop_pattern)
# # #     if dropped_prune:
# # #         st.info(f"Hidden {len(dropped_prune)} mostly-empty / unnamed / patterned columns.")

# # #     num_cols = df.select_dtypes(include="number").columns.tolist()
# # #     valid_y_cols, excluded_y_cols = [], []
# # #     for c in num_cols:
# # #         s = pd.to_numeric(df[c], errors="coerce").dropna()
# # #         (valid_y_cols if (not s.empty and s.abs().sum() != 0) else excluded_y_cols).append(c)

# # #     cat_cols = guess_label_cols(df, num_cols)
# # #     ghg_data, months_known = extract_ghg_data(raw)

# # #     # Try to auto-detect baseline from xlsx; fallback to 0.3522
# # #     auto_baseline = extract_baseline_from_raw(raw)

# # #     # ─────────────────────────────────────────────────────────────
# # #     # SIDEBAR
# # #     # ─────────────────────────────────────────────────────────────
# # #     with st.sidebar:
# # #         st.header("Controls")
# # #         x_col = st.selectbox("X-axis (categorical / time)", options=(cat_cols or [None]))
# # #         if valid_y_cols:
# # #             y_cols = st.multiselect("Y-axis (numeric)", options=valid_y_cols, default=valid_y_cols[:1])
# # #         else:
# # #             st.warning("No numeric columns with non-zero/nonnull values detected.")
# # #             y_cols = []

# # #         if y_cols:
# # #             y_numeric_all = df[y_cols].apply(pd.to_numeric, errors="coerce").fillna(0).abs()
# # #             df_for_filters = df.loc[(y_numeric_all.sum(axis=1) != 0)].copy()
# # #         else:
# # #             df_for_filters = df.copy()

# # #         st.markdown("**Filters**")
# # #         active_filters = {}
# # #         for c in cat_cols[:6]:
# # #             vals = df_for_filters[c].dropna().astype(str).unique().tolist()
# # #             if 1 < len(vals) <= 200:
# # #                 sel = st.multiselect(f"{c}", options=vals, default=[], placeholder="(all)")
# # #                 if sel:
# # #                     active_filters[c] = set(sel)

# # #         chart_type = st.radio("Chart type (generic)", ["Line", "Bar", "Area"], index=1)

# # #         st.markdown("---")
# # #         st.subheader("GHG Chart Options")
# # #         ghg_chart_type = st.radio("GHG Chart type", ["Line", "Bar", "Area"], index=0)
# # #         show_prediction = st.checkbox("Show Predicted Months", value=True)

# # #         st.markdown("---")
# # #         st.subheader("GEI Reference Lines")

# # #         # ── NEW: T line input ─────────────────────────────────────
# # #         t_line_value = st.number_input(
# # #             "Emission per ton of Equivalent product (T) — Green straight line",
# # #             value=0.315300059669252,
# # #             step=0.0001,
# # #             format="%.6f",
# # #             help="Green straight reference line for T value on the GEI chart",
# # #         )

# # #         # ── Existing: Target / Baseline input ────────────────────
# # #         default_bl = auto_baseline if auto_baseline else 0.3522
# # #         baseline_value = st.number_input(
# # #             "Baseline Target — Red dashed straight line",
# # #             value=default_bl,
# # #             step=0.001,
# # #             format="%.4f",
# # #             help="Red dashed reference line on the GEI Emission Intensity chart",
# # #         )

# # #     # ─────────────────────────────────────────────────────────────
# # #     # FILTERS
# # #     # ─────────────────────────────────────────────────────────────
# # #     filt = pd.Series(True, index=df.index)
# # #     for col, allowed in active_filters.items():
# # #         filt &= df[col].astype(str).isin(allowed)
# # #     df_f = df.loc[filt].copy()

# # #     if y_cols:
# # #         y_numeric_after = df_f[y_cols].apply(pd.to_numeric, errors="coerce")
# # #         nonzero_mask = (y_numeric_after.fillna(0).abs().sum(axis=1) != 0)
# # #         removed = (~nonzero_mask).sum()
# # #         if removed:
# # #             st.info(f"Removed {int(removed)} row(s) where selected Y columns were all zero/null.")
# # #         df_f = df_f.loc[nonzero_mask].reset_index(drop=True)

# # #     if excluded_y_cols:
# # #         st.info(f"Excluded numeric columns (all zero or all null): {', '.join(excluded_y_cols)}")
# # #     if df_f.empty:
# # #         st.warning("No rows left after filters. Adjust selection."); st.stop()

# # #     st.subheader("Data Table")
# # #     st.dataframe(arrow_safe(df_f), use_container_width=True, height=420)

# # #     # ─────────────────────────────────────────────────────────────
# # #     # ★ GHG DEDICATED SECTION ★
# # #     # ─────────────────────────────────────────────────────────────
# # #     any_ghg = any(bool(ghg_data[lbl]) for lbl in GHG_ROW_LABELS)

# # #     if any_ghg and months_known:
# # #         st.markdown("---")
# # #         st.header("📊 GHG Emission Analysis")

# # #         # Metric cards for known months (GEI row)
# # #         gei_label = "Emission per ton of Equivalent product"
# # #         col_info = st.columns(len(months_known))
# # #         for i, m in enumerate(months_known):
# # #             v = ghg_data[gei_label].get(m)
# # #             delta_val = round(v - baseline_value, 4) if v is not None else None
# # #             col_info[i].metric(
# # #                 label=m,
# # #                 value=f"{v:.4f}" if v is not None else "N/A",
# # #                 delta=f"{delta_val:+.4f} vs baseline {baseline_value:.4f}" if delta_val is not None else None,
# # #                 delta_color="inverse",
# # #             )

# # #         # ── 1. DEDICATED GEI CHART ─────────────────────────────────
# # #         st.markdown("---")
# # #         st.subheader("🎯 Emission Intensity per Ton — T Line (Green) · Baseline Target (Red) · Zone Fill")
# # #         if ghg_data.get(gei_label):
# # #             fig_gei = make_gei_chart(
# # #                 ghg_data=ghg_data,
# # #                 months_known=months_known,
# # #                 baseline_value=baseline_value,
# # #                 t_line_value=t_line_value,
# # #                 chart_type=ghg_chart_type,
# # #                 add_prediction=show_prediction,
# # #             )
# # #             st.plotly_chart(fig_gei, use_container_width=True)
# # #         else:
# # #             st.warning("No data for Emission per ton of Equivalent product.")

# # #         # ── 2. OTHER GHG CHARTS ────────────────────────────────────
# # #         OTHER_CHART_CONFIGS = [
# # #             {
# # #                 "label": "Total Direct and Indirect Emission",
# # #                 "y_axis": "tCO₂ eq",
# # #                 "color": "#E53935",
# # #             },
# # #             {
# # #                 "label": "Total Direct Emission (Scope 1)",
# # #                 "y_axis": "tCO₂ eq",
# # #                 "color": "#FF7043",
# # #             },
# # #             {
# # #                 "label": "Total Indirect Emission (Scope 2)",
# # #                 "y_axis": "tCO₂ eq",
# # #                 "color": "#7B1FA2",
# # #             },
# # #             {
# # #                 "label": "Total Equivalent Product for GHG Emission Intensity",
# # #                 "y_axis": "Tonnes",
# # #                 "color": "#2E7D32",
# # #             },
# # #         ]

# # #         st.markdown("---")
# # #         st.subheader("📈 Other GHG Metrics")
# # #         for cfg in OTHER_CHART_CONFIGS:
# # #             lbl = cfg["label"]
# # #             if not ghg_data.get(lbl):
# # #                 st.warning(f"No data found for: {lbl}")
# # #                 continue
# # #             fig = make_ghg_chart(
# # #                 ghg_data=ghg_data,
# # #                 label=lbl,
# # #                 months_known=months_known,
# # #                 chart_type=ghg_chart_type,
# # #                 y_axis_label=cfg["y_axis"],
# # #                 color=cfg["color"],
# # #             )
# # #             st.subheader(lbl)
# # #             st.plotly_chart(fig, use_container_width=True)

# # #         # ── 3. PREDICTION TABLE ────────────────────────────────────
# # #         if show_prediction:
# # #             remaining = get_remaining_fiscal_months(months_known)

# # #             st.markdown("---")
# # #             st.subheader("📅 Baseline Prediction — Remaining Months of This Fiscal Year")

# # #             if not remaining:
# # #                 st.info("No remaining months to predict — all fiscal year months have data.")
# # #             else:
# # #                 st.caption(
# # #                     f"Months with data: **{', '.join(months_known)}**  |  "
# # #                     f"Predicting: **{', '.join(remaining)}**"
# # #                 )
# # #                 pred_rows = []
# # #                 for lbl in GHG_ROW_LABELS:
# # #                     vals = ghg_data[lbl]
# # #                     if len(vals) < 2:
# # #                         continue
# # #                     mk = [m for m in FISCAL_YEAR_ORDER if m in vals]
# # #                     mv = [vals[m] for m in mk]
# # #                     try:
# # #                         pred_dict, slope, intercept = predict_remaining_months(mk, mv, FISCAL_YEAR_ORDER)
# # #                         for m in remaining:
# # #                             if m in pred_dict:
# # #                                 pred_rows.append({
# # #                                     "Metric": lbl,
# # #                                     "Month": m,
# # #                                     "Predicted Value": round(pred_dict[m], 6),
# # #                                 })
# # #                     except Exception:
# # #                         pass

# # #                 if pred_rows:
# # #                     pred_df = pd.DataFrame(pred_rows)
# # #                     pivot = pred_df.pivot(index="Metric", columns="Month", values="Predicted Value")
# # #                     pivot = pivot[[m for m in remaining if m in pivot.columns]]
# # #                     st.dataframe(pivot, use_container_width=True)
# # #                 else:
# # #                     st.info("Not enough data to generate predictions.")

# # #     # ─────────────────────────────────────────────────────────────
# # #     # GENERIC CHARTS
# # #     # ─────────────────────────────────────────────────────────────
# # #     if not y_cols:
# # #         st.warning("Pick at least one numeric Y column."); st.stop()

# # #     st.markdown("---")
# # #     st.header("📈 Generic Charts (Selected Columns)")

# # #     if x_col and x_col in df_f.columns:
# # #         x_vals = df_f[x_col].astype(str).fillna("")
# # #     else:
# # #         x_vals = pd.Series(range(len(df_f)), index=df_f.index, name="Index")
# # #         x_col = "Index"
# # #         df_f = df_f.assign(Index=x_vals)

# # #     for i, y in enumerate(y_cols):
# # #         primary_color = px.colors.qualitative.Plotly[i % len(px.colors.qualitative.Plotly)]
# # #         y_vals = pd.to_numeric(df_f[y], errors="coerce")
# # #         if y_vals.dropna().empty or y_vals.dropna().abs().sum() == 0:
# # #             st.info(f"Skipping '{y}' (all zero/null)."); continue
# # #         df_plot = pd.DataFrame({x_col: x_vals, y: y_vals.fillna(0)})

# # #         if chart_type == "Bar":
# # #             fig = px.bar(df_plot, x=x_col, y=y)
# # #             fig.update_traces(marker=dict(color=primary_color))
# # #         elif chart_type == "Line":
# # #             fig = px.line(df_plot, x=x_col, y=y)
# # #             fig.update_traces(line=dict(color=primary_color))
# # #         else:
# # #             fig = px.area(df_plot, x=x_col, y=y)
# # #             fig.update_traces(line=dict(color=primary_color),
# # #                               fillcolor=hex_to_rgba(primary_color, 0.25))

# # #         fig.update_layout(margin=dict(l=10, r=10, t=40, b=10), height=480, hovermode="x unified")
# # #         st.plotly_chart(fig, use_container_width=True)

# # #     st.download_button(
# # #         "Download cleaned/filtered data (CSV)",
# # #         df_f.to_csv(index=False).encode("utf-8"),
# # #         file_name="summary_cleaned_filtered.csv",
# # #         mime="text/csv",
# # #     )

# # # else:
# # #     st.info("Upload a file to begin.")


# # import streamlit as st
# # import pandas as pd
# # import numpy as np
# # import plotly.express as px
# # import plotly.graph_objects as go
# # import re
# # from scipy import stats

# # st.set_page_config(page_title="GHG Summary Analyzer", layout="wide")
# # st.title("GHG Summary Analyzer")
# # st.caption("Upload an Excel; we'll detect the Summary sheet, clean it, and build interactive charts.")

# # uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

# # # ─────────────────────────────────────────────────────────────────
# # # HELPERS
# # # ─────────────────────────────────────────────────────────────────
# # def load_excel(uploaded_file):
# #     for eng in ("openpyxl", "calamine"):
# #         try:
# #             return pd.ExcelFile(uploaded_file, engine=eng), eng
# #         except Exception:
# #             continue
# #     raise ImportError("No Excel engine available.")

# # def find_sheet(xls):
# #     if "Summary Sheet" in xls.sheet_names:
# #         return "Summary Sheet"
# #     for s in xls.sheet_names:
# #         if "summary" in s.lower():
# #             return s
# #     return xls.sheet_names[0]

# # def header_detect_clean(df_raw):
# #     raw = df_raw.dropna(how="all").dropna(axis=1, how="all").copy()
# #     header_idx = raw.head(min(15, len(raw))).notna().sum(axis=1).idxmax()
# #     df = raw.copy()
# #     df.columns = df.loc[header_idx].astype(str).str.strip()
# #     df = df.loc[header_idx + 1:].reset_index(drop=True)
# #     seen, cols = {}, []
# #     for c in df.columns:
# #         base = c if c and c != "nan" else "Unnamed"
# #         seen[base] = seen.get(base, 0) + 1
# #         cols.append(base if seen[base] == 1 else f"{base}_{seen[base]}")
# #     df.columns = cols
# #     for c in df.columns:
# #         parsed = pd.to_numeric(df[c], errors="coerce")
# #         if parsed.notna().mean() >= 0.6:
# #             df[c] = parsed
# #     return df

# # def prune_columns(df, drop_unnamed=True, null_threshold=0.90, drop_pattern=None):
# #     drop = []
# #     for c in df.columns:
# #         name = str(c).strip()
# #         if drop_unnamed and (name.lower().startswith("unnamed") or name == "" or name.lower() == "nan"):
# #             drop.append(c); continue
# #         if drop_pattern and re.match(drop_pattern, name):
# #             drop.append(c); continue
# #         if df[c].isna().mean() >= null_threshold:
# #             drop.append(c); continue
# #         if df[c].dtype == "object":
# #             s = df[c].astype(str).str.strip().replace("nan", "")
# #             if s.replace("", np.nan).dropna().empty:
# #                 drop.append(c); continue
# #     return df.drop(columns=drop), drop

# # def guess_label_cols(df, num_cols):
# #     return [c for c in df.columns if c not in num_cols]

# # def arrow_safe(df):
# #     out = df.copy()
# #     for c in out.columns:
# #         if out[c].dtype == "object":
# #             out[c] = out[c].astype("string")
# #     return out

# # def hex_to_rgba(hex_color, alpha):
# #     try:
# #         hc = str(hex_color).lstrip("#")
# #         if len(hc) != 6:
# #             return hex_color
# #         r, g, b = int(hc[0:2], 16), int(hc[2:4], 16), int(hc[4:6], 16)
# #         return f"rgba({r},{g},{b},{alpha})"
# #     except Exception:
# #         return hex_color

# # # ─────────────────────────────────────────────────────────────────
# # # GHG CONFIG
# # # ─────────────────────────────────────────────────────────────────
# # GHG_ROW_LABELS = {
# #     "Emission per ton of Equivalent product":               "T",
# #     "Total Direct and Indirect Emission":                   "S",
# #     "Total Direct Emission (Scope 1)":                      "Scope1",
# #     "Total Indirect Emission (Scope 2)":                    "Scope2",
# #     "Total Equivalent Product for GHG Emission Intensity":  "P",
# # }

# # KNOWN_MONTHS = ["January","February","March","April","May","June",
# #                 "July","August","September","October","November","December"]

# # FISCAL_YEAR_ORDER = ["April","May","June","July","August","September",
# #                      "October","November","December","January","February","March"]

# # OTHER_CHART_CONFIGS = [
# #     {"label": "Total Direct and Indirect Emission",                   "y_axis": "tCO₂ eq",  "color": "#E53935"},
# #     {"label": "Total Direct Emission (Scope 1)",                      "y_axis": "tCO₂ eq",  "color": "#FF7043"},
# #     {"label": "Total Indirect Emission (Scope 2)",                    "y_axis": "tCO₂ eq",  "color": "#7B1FA2"},
# #     {"label": "Total Equivalent Product for GHG Emission Intensity",  "y_axis": "Tonnes",   "color": "#2E7D32"},
# # ]

# # # ─────────────────────────────────────────────────────────────────
# # # GHG PARSER
# # # ─────────────────────────────────────────────────────────────────
# # def extract_ghg_data(raw):
# #     month_cols = {}
# #     for i, row in raw.iterrows():
# #         for col_idx, val in enumerate(row):
# #             vs = str(val).strip()
# #             for m in KNOWN_MONTHS:
# #                 if m.lower() == vs.lower():
# #                     month_cols[m] = col_idx
# #         if len(month_cols) >= 2:
# #             break

# #     results = {label: {} for label in GHG_ROW_LABELS}
# #     for i, row in raw.iterrows():
# #         row_vals = [str(v).strip() for v in row]
# #         for label in GHG_ROW_LABELS:
# #             if any(label.lower() in rv.lower() for rv in row_vals):
# #                 for month, col_idx in month_cols.items():
# #                     try:
# #                         v = pd.to_numeric(row.iloc[col_idx], errors="coerce")
# #                         if pd.notna(v):
# #                             results[label][month] = v
# #                     except Exception:
# #                         pass

# #     ordered_months = [m for m in FISCAL_YEAR_ORDER if m in month_cols]
# #     return results, ordered_months


# # def extract_baseline_from_raw(raw):
# #     for i, row in raw.iterrows():
# #         row_vals = [str(v).strip() for v in row]
# #         if any("emission per ton" in rv.lower() for rv in row_vals):
# #             numeric_vals = pd.to_numeric(pd.Series(row.tolist()), errors="coerce").dropna()
# #             if not numeric_vals.empty:
# #                 return float(numeric_vals.iloc[-1])
# #     return None


# # def get_remaining_fiscal_months(months_known):
# #     if not months_known:
# #         return []
# #     last_known = months_known[-1]
# #     if last_known not in FISCAL_YEAR_ORDER:
# #         return []
# #     last_idx = FISCAL_YEAR_ORDER.index(last_known)
# #     return FISCAL_YEAR_ORDER[last_idx + 1:]


# # def predict_remaining_months(months_known, values, all_months):
# #     if len(values) < 2:
# #         return {}, 0, 0
# #     x_known = np.array([all_months.index(m) for m in months_known if m in all_months])
# #     y_known = np.array(values)
# #     slope, intercept, *_ = stats.linregress(x_known, y_known)
# #     predictions = {}
# #     for m in all_months:
# #         if m not in months_known:
# #             predictions[m] = slope * all_months.index(m) + intercept
# #     return predictions, slope, intercept


# # # ─────────────────────────────────────────────────────────────────
# # # TIGHT Y-RANGE HELPER
# # # ─────────────────────────────────────────────────────────────────
# # def tight_y_range(values, pad_pct=0.25):
# #     clean = [v for v in values if v is not None and not np.isnan(v)]
# #     if not clean:
# #         return None
# #     y_min, y_max = min(clean), max(clean)
# #     span = y_max - y_min
# #     pad = span * pad_pct if span > 0 else abs(y_min) * 0.1 or 0.05
# #     return [max(0, y_min - pad), y_max + pad]


# # # ─────────────────────────────────────────────────────────────────
# # # GEI CHART — fixed y-axis clamped to data, month-filtered
# # # ─────────────────────────────────────────────────────────────────
# # def make_gei_chart(ghg_data, months_known, selected_months, baseline_value, t_line_value, add_prediction):
# #     label = "Emission per ton of Equivalent product"
# #     display_months = [m for m in FISCAL_YEAR_ORDER if m in selected_months]

# #     # Build predictions for ALL fiscal months
# #     pred_dict = {}
# #     if len(months_known) >= 2:
# #         known_vals = [ghg_data[label][m] for m in months_known if m in ghg_data[label]]
# #         result = predict_remaining_months(months_known, known_vals, FISCAL_YEAR_ORDER)
# #         if isinstance(result, tuple):
# #             pred_dict = result[0]

# #     actual_x, actual_y_vals = [], []
# #     pred_x, pred_y_vals = [], []

# #     for m in display_months:
# #         if m in ghg_data[label] and ghg_data[label][m] is not None:
# #             actual_x.append(m)
# #             actual_y_vals.append(ghg_data[label][m])
# #         elif add_prediction and m in pred_dict:
# #             pred_x.append(m)
# #             pred_y_vals.append(pred_dict[m])

# #     # Y-range: clamp tightly around all values including reference lines
# #     all_vals = actual_y_vals + pred_y_vals + [t_line_value, baseline_value]
# #     y_range = tight_y_range(all_vals, pad_pct=0.20)

# #     fig = go.Figure()

# #     # ── ZONE FILL between T and Target ──────────────────────────
# #     lower_val = min(t_line_value, baseline_value)
# #     upper_val = max(t_line_value, baseline_value)

# #     # invisible lower bound
# #     fig.add_trace(go.Scatter(
# #         x=display_months, y=[lower_val] * len(display_months),
# #         mode="lines", line=dict(width=0, color="rgba(0,0,0,0)"),
# #         showlegend=False, hoverinfo="skip",
# #     ))
# #     # upper bound fills down to lower
# #     fig.add_trace(go.Scatter(
# #         x=display_months, y=[upper_val] * len(display_months),
# #         mode="lines", line=dict(width=0, color="rgba(0,0,0,0)"),
# #         fill="tonexty", fillcolor="rgba(100,220,130,0.22)",
# #         showlegend=True, name="Zone: T ↔ Target", hoverinfo="skip",
# #     ))

# #     # ── ACTUAL GEI (solid bright green spline) ──────────────────
# #     if actual_x:
# #         fig.add_trace(go.Scatter(
# #             x=actual_x, y=actual_y_vals,
# #             name="Actual GEI",
# #             mode="lines+markers+text",
# #             line=dict(color="#00E676", width=3, shape="spline", smoothing=1.2),
# #             marker=dict(size=9, color="#00E676", line=dict(color="white", width=1.5)),
# #             text=[f"{v:.4f}" for v in actual_y_vals],
# #             textposition="top center",
# #             textfont=dict(color="#00E676", size=11),
# #         ))

# #     # ── PREDICTED GEI (dashed green spline) ─────────────────────
# #     if pred_x and add_prediction:
# #         if actual_x:
# #             fig.add_trace(go.Scatter(
# #                 x=[actual_x[-1], pred_x[0]], y=[actual_y_vals[-1], pred_y_vals[0]],
# #                 mode="lines", showlegend=False, hoverinfo="skip",
# #                 line=dict(color="#00E676", width=2, dash="dot", shape="spline"),
# #             ))
# #         fig.add_trace(go.Scatter(
# #             x=pred_x, y=pred_y_vals,
# #             name="Predicted GEI (Trend)",
# #             mode="lines+markers+text",
# #             line=dict(color="#00E676", width=2.5, dash="dot", shape="spline", smoothing=1.2),
# #             marker=dict(size=8, color="#00E676", symbol="diamond",
# #                         line=dict(color="white", width=1)),
# #             text=[f"{v:.4f}" for v in pred_y_vals],
# #             textposition="top center",
# #             textfont=dict(color="#00E676", size=10),
# #         ))

# #     # ── T LINE — solid green straight ───────────────────────────
# #     t_labels = [""] * len(display_months)
# #     if display_months:
# #         t_labels[-1] = f"T = {t_line_value:.4f}"
# #     fig.add_trace(go.Scatter(
# #         x=display_months, y=[t_line_value] * len(display_months),
# #         name=f"T Line = {t_line_value:.6f}",
# #         mode="lines+text",
# #         line=dict(color="#43A047", width=2.5, dash="solid"),
# #         text=t_labels,
# #         textposition="middle right",
# #         textfont=dict(color="#43A047", size=12, family="monospace"),
# #     ))

# #     # ── TARGET LINE — red dashed ─────────────────────────────────
# #     tgt_labels = [""] * len(display_months)
# #     if display_months:
# #         tgt_labels[-1] = f"Target = {baseline_value:.4f}"
# #     fig.add_trace(go.Scatter(
# #         x=display_months, y=[baseline_value] * len(display_months),
# #         name=f"Baseline Target = {baseline_value:.4f}",
# #         mode="lines+text",
# #         line=dict(color="#F44336", width=2.5, dash="dash"),
# #         text=tgt_labels,
# #         textposition="middle right",
# #         textfont=dict(color="#F44336", size=12, family="monospace"),
# #     ))

# #     fig.update_layout(
# #         title=dict(
# #             text="📌 GHG Emission Intensity — Actual GEI · T Line (Green) · Baseline Target (Red)",
# #             font=dict(size=15, color="#EEEEEE"),
# #         ),
# #         xaxis=dict(
# #             title="Month (Fiscal Year Apr → Mar)",
# #             tickfont=dict(size=12),
# #             gridcolor="rgba(200,200,200,0.12)",
# #             categoryorder="array",
# #             categoryarray=display_months,
# #         ),
# #         yaxis=dict(
# #             title=f"GEI (tCO₂ eq/ton)   ┃   T: {t_line_value:.4f}   ┃   Baseline: {baseline_value:.4f}",
# #             range=y_range,
# #             gridcolor="rgba(200,200,200,0.12)",
# #             zeroline=False,
# #             tickfont=dict(size=11),
# #         ),
# #         height=540,
# #         margin=dict(l=10, r=10, t=70, b=10),
# #         hovermode="x unified",
# #         legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
# #                     font=dict(size=11)),
# #         plot_bgcolor="rgba(15,20,30,0.9)",
# #         paper_bgcolor="rgba(0,0,0,0)",
# #     )
# #     return fig


# # # ─────────────────────────────────────────────────────────────────
# # # OTHER GHG CHART — month-filtered, tight y-axis
# # # ─────────────────────────────────────────────────────────────────
# # def make_ghg_chart(ghg_data, label, selected_months, chart_type, y_axis_label, color):
# #     display_months = [m for m in FISCAL_YEAR_ORDER if m in selected_months]
# #     actual_y = [ghg_data[label].get(m, None) for m in display_months]
# #     rgba_fill = hex_to_rgba(color, 0.22)
# #     y_range = tight_y_range(actual_y, pad_pct=0.25)

# #     fig = go.Figure()
# #     if chart_type == "Bar":
# #         fig.add_trace(go.Bar(
# #             x=display_months,
# #             y=[v if v is not None else 0 for v in actual_y],
# #             name="Actual",
# #             marker_color=[color if v is not None else "rgba(0,0,0,0)" for v in actual_y],
# #             text=[f"{v:,.2f}" if v is not None else "" for v in actual_y],
# #             textposition="outside",
# #         ))
# #     elif chart_type == "Line":
# #         fig.add_trace(go.Scatter(
# #             x=display_months, y=actual_y, name="Actual",
# #             mode="lines+markers+text",
# #             line=dict(color=color, width=2.5, shape="spline", smoothing=1.0),
# #             marker=dict(size=8),
# #             text=[f"{v:,.2f}" if v is not None else "" for v in actual_y],
# #             textposition="top center",
# #         ))
# #     else:  # Area
# #         fig.add_trace(go.Scatter(
# #             x=display_months, y=actual_y, name="Actual",
# #             mode="lines+markers", fill="tozeroy",
# #             line=dict(color=color, width=2.5, shape="spline", smoothing=1.0),
# #             fillcolor=rgba_fill,
# #             marker=dict(size=8),
# #         ))

# #     fig.update_layout(
# #         xaxis=dict(
# #             title="Month (Fiscal Year Apr → Mar)",
# #             tickfont=dict(size=12),
# #             categoryorder="array",
# #             categoryarray=display_months,
# #         ),
# #         yaxis=dict(
# #             title=y_axis_label,
# #             range=y_range,
# #             zeroline=False,
# #         ),
# #         height=420,
# #         margin=dict(l=10, r=10, t=40, b=10),
# #         hovermode="x unified",
# #         legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
# #     )
# #     return fig


# # # ─────────────────────────────────────────────────────────────────
# # # MAIN APP
# # # ─────────────────────────────────────────────────────────────────
# # if uploaded:
# #     try:
# #         xls, engine_used = load_excel(uploaded)
# #         st.caption(f"Excel engine: **{engine_used}**")
# #         sheet = find_sheet(xls)
# #         st.success(f"Using sheet: **{sheet}**")
# #         raw = pd.read_excel(uploaded, sheet_name=sheet, engine=engine_used, header=None)
# #         df = header_detect_clean(raw)
# #     except Exception as e:
# #         st.exception(e); st.stop()

# #     drop_pattern = r"^0\.0_.*"
# #     df, dropped_prune = prune_columns(df, drop_unnamed=True, null_threshold=0.90, drop_pattern=drop_pattern)
# #     if dropped_prune:
# #         st.info(f"Hidden {len(dropped_prune)} mostly-empty / unnamed / patterned columns.")

# #     num_cols = df.select_dtypes(include="number").columns.tolist()
# #     valid_y_cols, excluded_y_cols = [], []
# #     for c in num_cols:
# #         s = pd.to_numeric(df[c], errors="coerce").dropna()
# #         (valid_y_cols if (not s.empty and s.abs().sum() != 0) else excluded_y_cols).append(c)

# #     cat_cols = guess_label_cols(df, num_cols)
# #     ghg_data, months_known = extract_ghg_data(raw)
# #     auto_baseline = extract_baseline_from_raw(raw)

# #     # ─────────────────────────────────────────────────────────────
# #     # SIDEBAR
# #     # ─────────────────────────────────────────────────────────────
# #     with st.sidebar:
# #         st.header("Controls")

# #         # ── Month multiselect — controls ALL GHG charts ──────────
# #         st.subheader("📅 Month Filter")
# #         st.caption("Controls all GHG charts & metric cards")
# #         all_available_months = [m for m in FISCAL_YEAR_ORDER
# #                                  if any(m in ghg_data[lbl] for lbl in GHG_ROW_LABELS)]
# #         default_months = months_known if months_known else all_available_months
# #         selected_months = st.multiselect(
# #             "Select Months to Display",
# #             options=FISCAL_YEAR_ORDER,
# #             default=default_months,
# #             placeholder="Choose months…",
# #         )
# #         if not selected_months:
# #             selected_months = default_months

# #         st.markdown("---")

# #         # ── Generic chart controls ───────────────────────────────
# #         x_col = st.selectbox("X-axis (categorical / time)", options=(cat_cols or [None]))
# #         if valid_y_cols:
# #             y_cols = st.multiselect("Y-axis (numeric)", options=valid_y_cols, default=valid_y_cols[:1])
# #         else:
# #             st.warning("No numeric columns detected.")
# #             y_cols = []

# #         if y_cols:
# #             y_numeric_all = df[y_cols].apply(pd.to_numeric, errors="coerce").fillna(0).abs()
# #             df_for_filters = df.loc[(y_numeric_all.sum(axis=1) != 0)].copy()
# #         else:
# #             df_for_filters = df.copy()

# #         st.markdown("**Row Filters**")
# #         active_filters = {}
# #         for c in cat_cols[:6]:
# #             vals = df_for_filters[c].dropna().astype(str).unique().tolist()
# #             if 1 < len(vals) <= 200:
# #                 sel = st.multiselect(f"{c}", options=vals, default=[], placeholder="(all)")
# #                 if sel:
# #                     active_filters[c] = set(sel)

# #         chart_type = st.radio("Chart type (generic)", ["Line", "Bar", "Area"], index=1)

# #         st.markdown("---")
# #         st.subheader("GHG Chart Options")
# #         ghg_chart_type = st.radio("GHG Chart type", ["Line", "Bar", "Area"], index=0)
# #         show_prediction = st.checkbox("Show Predicted Months", value=True)

# #         st.markdown("---")
# #         st.subheader("GEI Reference Lines")

# #         t_line_value = st.number_input(
# #             "Emission per ton of Equivalent product (T)\n— Green straight line",
# #             value=0.315300059669252,
# #             step=0.0001,
# #             format="%.6f",
# #             help="Solid green straight reference line on GEI chart",
# #         )

# #         default_bl = auto_baseline if auto_baseline else 0.3522
# #         baseline_value = st.number_input(
# #             "Baseline Target — Red dashed straight line",
# #             value=0.3522,
# #             step=0.0001,
# #             format="%.4f",
# #             help="Red dashed reference line on GEI chart",
# #         )

# #     # ─────────────────────────────────────────────────────────────
# #     # ROW FILTERS
# #     # ─────────────────────────────────────────────────────────────
# #     filt = pd.Series(True, index=df.index)
# #     for col, allowed in active_filters.items():
# #         filt &= df[col].astype(str).isin(allowed)
# #     df_f = df.loc[filt].copy()

# #     if y_cols:
# #         y_numeric_after = df_f[y_cols].apply(pd.to_numeric, errors="coerce")
# #         nonzero_mask = (y_numeric_after.fillna(0).abs().sum(axis=1) != 0)
# #         removed = (~nonzero_mask).sum()
# #         if removed:
# #             st.info(f"Removed {int(removed)} row(s) where all selected Y = zero/null.")
# #         df_f = df_f.loc[nonzero_mask].reset_index(drop=True)

# #     if excluded_y_cols:
# #         st.info(f"Excluded numeric columns (all zero / null): {', '.join(excluded_y_cols)}")
# #     if df_f.empty:
# #         st.warning("No rows left after filters. Adjust selection."); st.stop()

# #     st.subheader("Data Table")
# #     st.dataframe(arrow_safe(df_f), use_container_width=True, height=360)

# #     # ─────────────────────────────────────────────────────────────
# #     # ★ GHG SECTION ★
# #     # ─────────────────────────────────────────────────────────────
# #     any_ghg = any(bool(ghg_data[lbl]) for lbl in GHG_ROW_LABELS)

# #     if any_ghg and months_known:
# #         st.markdown("---")
# #         st.header("📊 GHG Emission Analysis")

# #         # Month filter badge
# #         display_month_list = [m for m in FISCAL_YEAR_ORDER if m in selected_months]
# #         st.caption(f"📅 Showing: **{', '.join(display_month_list) if display_month_list else 'None selected'}**")

# #         # ── Metric cards (filtered months only) ──────────────────
# #         gei_label = "Emission per ton of Equivalent product"
# #         months_to_card = [m for m in display_month_list if m in months_known]
# #         if months_to_card:
# #             ncols = min(len(months_to_card), 6)
# #             cols_cards = st.columns(ncols)
# #             for i, m in enumerate(months_to_card):
# #                 v = ghg_data[gei_label].get(m)
# #                 delta_val = round(v - baseline_value, 4) if v is not None else None
# #                 cols_cards[i % ncols].metric(
# #                     label=m,
# #                     value=f"{v:.4f}" if v is not None else "N/A",
# #                     delta=f"{delta_val:+.4f} vs baseline" if delta_val is not None else None,
# #                     delta_color="inverse",
# #                 )

# #         # ── 1. GEI CHART ──────────────────────────────────────────
# #         st.markdown("---")
# #         st.subheader("🎯 Emission Intensity per Ton — GEI")
# #         st.caption(
# #             f"🟢 Green solid line = T ({t_line_value:.6f})   |   "
# #             f"🔴 Red dashed line = Baseline Target ({baseline_value:.4f})   |   "
# #             f"Shaded band = zone between T and Target"
# #         )

# #         if ghg_data.get(gei_label):
# #             fig_gei = make_gei_chart(
# #                 ghg_data=ghg_data,
# #                 months_known=months_known,
# #                 selected_months=selected_months,
# #                 baseline_value=baseline_value,
# #                 t_line_value=t_line_value,
# #                 add_prediction=show_prediction,
# #             )
# #             st.plotly_chart(fig_gei, use_container_width=True)
# #         else:
# #             st.warning("No GEI data found in sheet.")

# #         # ── 2. OTHER GHG CHARTS ────────────────────────────────────
# #         st.markdown("---")
# #         st.subheader("📈 Other GHG Metrics")

# #         for cfg in OTHER_CHART_CONFIGS:
# #             lbl = cfg["label"]
# #             if not ghg_data.get(lbl):
# #                 st.warning(f"No data found for: {lbl}")
# #                 continue
# #             available = {m: ghg_data[lbl][m] for m in selected_months if m in ghg_data[lbl]}
# #             if not available:
# #                 st.info(f"No data for **{lbl}** in the selected months.")
# #                 continue
# #             fig = make_ghg_chart(
# #                 ghg_data=ghg_data,
# #                 label=lbl,
# #                 selected_months=selected_months,
# #                 chart_type=ghg_chart_type,
# #                 y_axis_label=cfg["y_axis"],
# #                 color=cfg["color"],
# #             )
# #             st.subheader(lbl)
# #             st.plotly_chart(fig, use_container_width=True)

# #         # ── 3. PREDICTION TABLE ────────────────────────────────────
# #         if show_prediction:
# #             remaining = get_remaining_fiscal_months(months_known)
# #             st.markdown("---")
# #             st.subheader("📅 Baseline Prediction — Remaining Fiscal Months")

# #             if not remaining:
# #                 st.info("No remaining months to predict — full fiscal year has data.")
# #             else:
# #                 st.caption(
# #                     f"Known: **{', '.join(months_known)}**   |   "
# #                     f"Predicting: **{', '.join(remaining)}**"
# #                 )
# #                 pred_rows = []
# #                 for lbl in GHG_ROW_LABELS:
# #                     vals = ghg_data[lbl]
# #                     if len(vals) < 2:
# #                         continue
# #                     mk = [m for m in FISCAL_YEAR_ORDER if m in vals]
# #                     mv = [vals[m] for m in mk]
# #                     try:
# #                         pred_dict, _, _ = predict_remaining_months(mk, mv, FISCAL_YEAR_ORDER)
# #                         for m in remaining:
# #                             if m in pred_dict:
# #                                 pred_rows.append({
# #                                     "Metric": lbl,
# #                                     "Month": m,
# #                                     "Predicted Value": round(pred_dict[m], 6),
# #                                 })
# #                     except Exception:
# #                         pass

# #                 if pred_rows:
# #                     pred_df = pd.DataFrame(pred_rows)
# #                     pivot = pred_df.pivot(index="Metric", columns="Month", values="Predicted Value")
# #                     pivot = pivot[[m for m in remaining if m in pivot.columns]]
# #                     st.dataframe(pivot, use_container_width=True)
# #                 else:
# #                     st.info("Not enough data to generate predictions.")

# #     # ─────────────────────────────────────────────────────────────
# #     # GENERIC CHARTS
# #     # ─────────────────────────────────────────────────────────────
# #     # if not y_cols:
# #     #     st.warning("Pick at least one numeric Y column in the sidebar."); st.stop()

# #     # st.markdown("---")
# #     # st.header("📈 Generic Charts (Selected Columns)")

# #     # if x_col and x_col in df_f.columns:
# #     #     x_vals = df_f[x_col].astype(str).fillna("")
# #     # else:
# #     #     x_vals = pd.Series(range(len(df_f)), index=df_f.index, name="Index")
# #     #     x_col = "Index"
# #     #     df_f = df_f.assign(Index=x_vals)

# #     # for i, y in enumerate(y_cols):
# #     #     primary_color = px.colors.qualitative.Plotly[i % len(px.colors.qualitative.Plotly)]
# #     #     y_vals = pd.to_numeric(df_f[y], errors="coerce")
# #     #     if y_vals.dropna().empty or y_vals.dropna().abs().sum() == 0:
# #     #         st.info(f"Skipping '{y}' (all zero/null)."); continue
# #     #     df_plot = pd.DataFrame({x_col: x_vals, y: y_vals.fillna(0)})

# #     #     if chart_type == "Bar":
# #     #         fig = px.bar(df_plot, x=x_col, y=y)
# #     #         fig.update_traces(marker=dict(color=primary_color))
# #     #     elif chart_type == "Line":
# #     #         fig = px.line(df_plot, x=x_col, y=y)
# #     #         fig.update_traces(line=dict(color=primary_color))
# #     #     else:
# #     #         fig = px.area(df_plot, x=x_col, y=y)
# #     #         fig.update_traces(line=dict(color=primary_color),
# #     #                           fillcolor=hex_to_rgba(primary_color, 0.25))

# #     #     fig.update_layout(margin=dict(l=10, r=10, t=40, b=10), height=450, hovermode="x unified")
# #     #     st.plotly_chart(fig, use_container_width=True)

# #     # st.download_button(
# #     #     "⬇️ Download cleaned/filtered data (CSV)",
# #     #     df_f.to_csv(index=False).encode("utf-8"),
# #     #     file_name="ghg_cleaned_filtered.csv",
# #     #     mime="text/csv",
# #     # )

# # else:
# #     st.info("⬆️ Upload an Excel file to begin.")

# import streamlit as st
# import pandas as pd
# import numpy as np
# import plotly.express as px
# import plotly.graph_objects as go
# import re
# from scipy import stats

# st.set_page_config(page_title="GHG Summary Analyzer", layout="wide")
# st.title("GHG Summary Analyzer")
# st.caption("Upload an Excel; we'll detect the Summary sheet, clean it, and build interactive charts.")

# uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

# # ─────────────────────────────────────────────────────────────────
# # HELPERS (unchanged)
# # ─────────────────────────────────────────────────────────────────
# def load_excel(uploaded_file):
#     for eng in ("openpyxl", "calamine"):
#         try:
#             return pd.ExcelFile(uploaded_file, engine=eng), eng
#         except Exception:
#             continue
#     raise ImportError("No Excel engine available.")

# def find_sheet(xls):
#     if "Summary Sheet" in xls.sheet_names:
#         return "Summary Sheet"
#     for s in xls.sheet_names:
#         if "summary" in s.lower():
#             return s
#     return xls.sheet_names[0]

# def header_detect_clean(df_raw):
#     raw = df_raw.dropna(how="all").dropna(axis=1, how="all").copy()
#     header_idx = raw.head(min(15, len(raw))).notna().sum(axis=1).idxmax()
#     df = raw.copy()
#     df.columns = df.loc[header_idx].astype(str).str.strip()
#     df = df.loc[header_idx + 1:].reset_index(drop=True)
#     seen, cols = {}, []
#     for c in df.columns:
#         base = c if c and c != "nan" else "Unnamed"
#         seen[base] = seen.get(base, 0) + 1
#         cols.append(base if seen[base] == 1 else f"{base}_{seen[base]}")
#     df.columns = cols
#     for c in df.columns:
#         parsed = pd.to_numeric(df[c], errors="coerce")
#         if parsed.notna().mean() >= 0.6:
#             df[c] = parsed
#     return df

# def prune_columns(df, drop_unnamed=True, null_threshold=0.90, drop_pattern=None):
#     drop = []
#     for c in df.columns:
#         name = str(c).strip()
#         if drop_unnamed and (name.lower().startswith("unnamed") or name == "" or name.lower() == "nan"):
#             drop.append(c); continue
#         if drop_pattern and re.match(drop_pattern, name):
#             drop.append(c); continue
#         if df[c].isna().mean() >= null_threshold:
#             drop.append(c); continue
#         if df[c].dtype == "object":
#             s = df[c].astype(str).str.strip().replace("nan", "")
#             if s.replace("", np.nan).dropna().empty:
#                 drop.append(c); continue
#     return df.drop(columns=drop), drop

# def guess_label_cols(df, num_cols):
#     return [c for c in df.columns if c not in num_cols]

# def arrow_safe(df):
#     out = df.copy()
#     for c in out.columns:
#         if out[c].dtype == "object":
#             out[c] = out[c].astype("string")
#     return out

# def hex_to_rgba(hex_color, alpha):
#     try:
#         hc = str(hex_color).lstrip("#")
#         if len(hc) != 6:
#             return hex_color
#         r, g, b = int(hc[0:2], 16), int(hc[2:4], 16), int(hc[4:6], 16)
#         return f"rgba({r},{g},{b},{alpha})"
#     except Exception:
#         return hex_color

# # ─────────────────────────────────────────────────────────────────
# # GHG CONFIG
# # ─────────────────────────────────────────────────────────────────
# GHG_ROW_LABELS = {
#     "Emission per ton of Equivalent product":               "T",
#     "Total Direct and Indirect Emission":                   "S",
#     "Total Direct Emission (Scope 1)":                      "Scope1",
#     "Total Indirect Emission (Scope 2)":                    "Scope2",
#     "Total Equivalent Product for GHG Emission Intensity":  "P",
# }

# KNOWN_MONTHS = ["January","February","March","April","May","June",
#                 "July","August","September","October","November","December"]

# FISCAL_YEAR_ORDER = ["April","May","June","July","August","September",
#                      "October","November","December","January","February","March"]

# OTHER_CHART_CONFIGS = [
#     {"label": "Total Direct and Indirect Emission",                   "y_axis": "tCO₂ eq",  "color": "#E53935"},
#     {"label": "Total Direct Emission (Scope 1)",                      "y_axis": "tCO₂ eq",  "color": "#FF7043"},
#     {"label": "Total Indirect Emission (Scope 2)",                    "y_axis": "tCO₂ eq",  "color": "#7B1FA2"},
#     {"label": "Total Equivalent Product for GHG Emission Intensity",  "y_axis": "Tonnes",   "color": "#2E7D32"},
# ]

# # ─────────────────────────────────────────────────────────────────
# # GHG PARSER (unchanged)
# # ─────────────────────────────────────────────────────────────────
# def extract_ghg_data(raw):
#     month_cols = {}
#     for i, row in raw.iterrows():
#         for col_idx, val in enumerate(row):
#             vs = str(val).strip()
#             for m in KNOWN_MONTHS:
#                 if m.lower() == vs.lower():
#                     month_cols[m] = col_idx
#         if len(month_cols) >= 2:
#             break

#     results = {label: {} for label in GHG_ROW_LABELS}
#     for i, row in raw.iterrows():
#         row_vals = [str(v).strip() for v in row]
#         for label in GHG_ROW_LABELS:
#             if any(label.lower() in rv.lower() for rv in row_vals):
#                 for month, col_idx in month_cols.items():
#                     try:
#                         v = pd.to_numeric(row.iloc[col_idx], errors="coerce")
#                         if pd.notna(v):
#                             results[label][month] = v
#                     except Exception:
#                         pass

#     ordered_months = [m for m in FISCAL_YEAR_ORDER if m in month_cols]
#     return results, ordered_months

# def extract_baseline_from_raw(raw):
#     for i, row in raw.iterrows():
#         row_vals = [str(v).strip() for v in row]
#         if any("emission per ton" in rv.lower() for rv in row_vals):
#             numeric_vals = pd.to_numeric(pd.Series(row.tolist()), errors="coerce").dropna()
#             if not numeric_vals.empty:
#                 return float(numeric_vals.iloc[-1])
#     return None

# def get_remaining_fiscal_months(months_known):
#     if not months_known:
#         return []
#     last_known = months_known[-1]
#     if last_known not in FISCAL_YEAR_ORDER:
#         return []
#     last_idx = FISCAL_YEAR_ORDER.index(last_known)
#     return FISCAL_YEAR_ORDER[last_idx + 1:]

# def predict_remaining_months(months_known, values, all_months):
#     if len(values) < 2:
#         return {}, 0, 0
#     x_known = np.array([all_months.index(m) for m in months_known if m in all_months])
#     y_known = np.array(values)
#     slope, intercept, *_ = stats.linregress(x_known, y_known)
#     predictions = {}
#     for m in all_months:
#         if m not in months_known:
#             predictions[m] = slope * all_months.index(m) + intercept
#     return predictions, slope, intercept

# def tight_y_range(values, pad_pct=0.25):
#     clean = [v for v in values if v is not None and not np.isnan(v)]
#     if not clean:
#         return None
#     y_min, y_max = min(clean), max(clean)
#     span = y_max - y_min
#     pad = span * pad_pct if span > 0 else abs(y_min) * 0.1 or 0.05
#     return [max(0, y_min - pad), y_max + pad]

# # ─────────────────────────────────────────────────────────────────
# # GEI CHART (unchanged - your original)
# # ─────────────────────────────────────────────────────────────────
# def make_gei_chart(ghg_data, months_known, selected_months, baseline_value, t_line_value, add_prediction):
#     label = "Emission per ton of Equivalent product"
#     display_months = [m for m in FISCAL_YEAR_ORDER if m in selected_months]

#     pred_dict = {}
#     if len(months_known) >= 2:
#         known_vals = [ghg_data[label][m] for m in months_known if m in ghg_data[label]]
#         result = predict_remaining_months(months_known, known_vals, FISCAL_YEAR_ORDER)
#         if isinstance(result, tuple):
#             pred_dict = result[0]

#     actual_x, actual_y_vals = [], []
#     pred_x, pred_y_vals = [], []

#     for m in display_months:
#         if m in ghg_data[label] and ghg_data[label][m] is not None:
#             actual_x.append(m)
#             actual_y_vals.append(ghg_data[label][m])
#         elif add_prediction and m in pred_dict:
#             pred_x.append(m)
#             pred_y_vals.append(pred_dict[m])

#     all_vals = actual_y_vals + pred_y_vals + [t_line_value, baseline_value]
#     y_range = tight_y_range(all_vals, pad_pct=0.20)

#     fig = go.Figure()

#     # Zone Fill
#     lower_val = min(t_line_value, baseline_value)
#     upper_val = max(t_line_value, baseline_value)
#     fig.add_trace(go.Scatter(x=display_months, y=[lower_val] * len(display_months), mode="lines", line=dict(width=0), showlegend=False))
#     fig.add_trace(go.Scatter(x=display_months, y=[upper_val] * len(display_months), mode="lines", line=dict(width=0),
#                              fill="tonexty", fillcolor="rgba(100,220,130,0.22)", name="Zone: T ↔ Target"))

#     # Actual GEI
#     if actual_x:
#         fig.add_trace(go.Scatter(
#             x=actual_x, y=actual_y_vals, name="Actual GEI",
#             mode="lines+markers+text", line=dict(color="#00E676", width=3, shape="spline", smoothing=1.2),
#             marker=dict(size=9, color="#00E676", line=dict(color="white", width=1.5)),
#             text=[f"{v:.4f}" for v in actual_y_vals], textposition="top center"
#         ))

#     # Predicted GEI
#     if pred_x and add_prediction:
#         if actual_x:
#             fig.add_trace(go.Scatter(x=[actual_x[-1], pred_x[0]], y=[actual_y_vals[-1], pred_y_vals[0]],
#                                      mode="lines", showlegend=False, line=dict(color="#00E676", width=2, dash="dot")))
#         fig.add_trace(go.Scatter(
#             x=pred_x, y=pred_y_vals, name="Predicted GEI (Trend)",
#             mode="lines+markers+text", line=dict(color="#00E676", width=2.5, dash="dot", shape="spline", smoothing=1.2),
#             marker=dict(size=8, color="#00E676", symbol="diamond")
#         ))

#     # T Line
#     t_labels = [""] * len(display_months)
#     if display_months:
#         t_labels[-1] = f"T = {t_line_value:.4f}"
#     fig.add_trace(go.Scatter(x=display_months, y=[t_line_value] * len(display_months),
#                              name=f"T Line = {t_line_value:.6f}", mode="lines+text",
#                              line=dict(color="#43A047", width=2.5),
#                              text=t_labels, textposition="middle right"))

#     # Baseline Target Line (Static 0.3552 in display, but value from input)
#     tgt_labels = [""] * len(display_months)
#     if display_months:
#         tgt_labels[-1] = f"Target = {baseline_value:.4f}"
#     fig.add_trace(go.Scatter(x=display_months, y=[baseline_value] * len(display_months),
#                              name=f"Baseline Target = {baseline_value:.4f}", mode="lines+text",
#                              line=dict(color="#F44336", width=2.5, dash="dash"),
#                              text=tgt_labels, textposition="middle right"))

#     fig.update_layout(
#         title="📌 GHG Emission Intensity — Actual GEI · T Line (Green) · Baseline Target (Red)",
#         xaxis=dict(title="Month (Fiscal Year Apr → Mar)", categoryorder="array", categoryarray=display_months),
#         yaxis=dict(title=f"GEI (tCO₂ eq/ton)   ┃   T: {t_line_value:.4f}   ┃   Baseline: {baseline_value:.4f}", range=y_range),
#         height=540,
#         hovermode="x unified",
#         legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
#         plot_bgcolor="rgba(15,20,30,0.9)",
#         paper_bgcolor="rgba(0,0,0,0)",
#     )
#     return fig

# # make_ghg_chart function remains unchanged (your original)
# def make_ghg_chart(ghg_data, label, selected_months, chart_type, y_axis_label, color):
#     display_months = [m for m in FISCAL_YEAR_ORDER if m in selected_months]
#     actual_y = [ghg_data[label].get(m, None) for m in display_months]
#     rgba_fill = hex_to_rgba(color, 0.22)
#     y_range = tight_y_range(actual_y, pad_pct=0.25)

#     fig = go.Figure()
#     if chart_type == "Bar":
#         fig.add_trace(go.Bar(x=display_months, y=[v if v is not None else 0 for v in actual_y],
#                              marker_color=[color if v is not None else "rgba(0,0,0,0)" for v in actual_y],
#                              text=[f"{v:,.2f}" if v is not None else "" for v in actual_y], textposition="outside"))
#     elif chart_type == "Line":
#         fig.add_trace(go.Scatter(x=display_months, y=actual_y, mode="lines+markers+text",
#                                  line=dict(color=color, width=2.5, shape="spline"), text=[f"{v:,.2f}" if v else "" for v in actual_y]))
#     else:
#         fig.add_trace(go.Scatter(x=display_months, y=actual_y, mode="lines+markers", fill="tozeroy",
#                                  line=dict(color=color, width=2.5, shape="spline"), fillcolor=rgba_fill))

#     fig.update_layout(xaxis=dict(title="Month (Fiscal Year Apr → Mar)", categoryorder="array", categoryarray=display_months),
#                       yaxis=dict(title=y_axis_label, range=y_range),
#                       height=420, hovermode="x unified")
#     return fig

# # ─────────────────────────────────────────────────────────────────
# # MAIN APP
# # ─────────────────────────────────────────────────────────────────
# if uploaded:
#     try:
#         xls, engine_used = load_excel(uploaded)
#         st.caption(f"Excel engine: **{engine_used}**")
#         sheet = find_sheet(xls)
#         st.success(f"Using sheet: **{sheet}**")
#         raw = pd.read_excel(uploaded, sheet_name=sheet, engine=engine_used, header=None)
#         df = header_detect_clean(raw)
#     except Exception as e:
#         st.exception(e)
#         st.stop()

#     drop_pattern = r"^0\.0_.*"
#     df, dropped_prune = prune_columns(df, drop_unnamed=True, null_threshold=0.90, drop_pattern=drop_pattern)
#     if dropped_prune:
#         st.info(f"Hidden {len(dropped_prune)} mostly-empty / unnamed / patterned columns.")

#     num_cols = df.select_dtypes(include="number").columns.tolist()
#     valid_y_cols, excluded_y_cols = [], []
#     for c in num_cols:
#         s = pd.to_numeric(df[c], errors="coerce").dropna()
#         (valid_y_cols if (not s.empty and s.abs().sum() != 0) else excluded_y_cols).append(c)

#     cat_cols = guess_label_cols(df, num_cols)
#     ghg_data, months_known = extract_ghg_data(raw)
#     auto_baseline = extract_baseline_from_raw(raw)

#     # ─────────────────────────────────────────────────────────────
#     # SIDEBAR - All options included
#     # ─────────────────────────────────────────────────────────────
#     with st.sidebar:
#         st.header("Controls")

#         st.subheader("📅 Month Filter")
#         all_available_months = [m for m in FISCAL_YEAR_ORDER if any(m in ghg_data[lbl] for lbl in GHG_ROW_LABELS)]
#         default_months = months_known if months_known else all_available_months
#         selected_months = st.multiselect(
#             "Select Months to Display",
#             options=FISCAL_YEAR_ORDER,
#             default=default_months,
#             placeholder="Choose months…"
#         )
#         if not selected_months:
#             selected_months = default_months

#         st.markdown("---")

#         x_col = st.selectbox("X-axis (categorical / time)", options=(cat_cols or [None]))
#         if valid_y_cols:
#             y_cols = st.multiselect("Y-axis (numeric)", options=valid_y_cols, default=valid_y_cols[:1])
#         else:
#             st.warning("No numeric columns detected.")
#             y_cols = []

#         if y_cols:
#             y_numeric_all = df[y_cols].apply(pd.to_numeric, errors="coerce").fillna(0).abs()
#             df_for_filters = df.loc[(y_numeric_all.sum(axis=1) != 0)].copy()
#         else:
#             df_for_filters = df.copy()

#         st.markdown("**Row Filters**")
#         active_filters = {}
#         for c in cat_cols[:6]:
#             vals = df_for_filters[c].dropna().astype(str).unique().tolist()
#             if 1 < len(vals) <= 200:
#                 sel = st.multiselect(f"{c}", options=vals, default=[], placeholder="(all)")
#                 if sel:
#                     active_filters[c] = set(sel)

#         chart_type = st.radio("Chart type (generic)", ["Line", "Bar", "Area"], index=1)

#         st.markdown("---")
#         st.subheader("GHG Chart Options")
#         ghg_chart_type = st.radio("GHG Chart type", ["Line", "Bar", "Area"], index=0)
#         show_prediction = st.checkbox("Show Predicted Months", value=True)

#         st.markdown("---")
#         st.subheader("GEI Reference Lines")

#         t_line_value = st.number_input(
#             "Emission per ton of Equivalent product (T) — Green straight line",
#             value=0.315300059669252,
#             step=0.0001,
#             format="%.6f"
#         )

#         # Baseline Target is STATIC 0.3552 as requested
#         baseline_value = 0.3552   # ← STATIC VALUE
#         st.info(f"**Baseline Target (Red)** is fixed at **{baseline_value:.4f}**")

#     # Row Filters
#     filt = pd.Series(True, index=df.index)
#     for col, allowed in active_filters.items():
#         filt &= df[col].astype(str).isin(allowed)
#     df_f = df.loc[filt].copy()

#     if y_cols:
#         y_numeric_after = df_f[y_cols].apply(pd.to_numeric, errors="coerce")
#         nonzero_mask = (y_numeric_after.fillna(0).abs().sum(axis=1) != 0)
#         if (~nonzero_mask).sum() > 0:
#             st.info(f"Removed {int((~nonzero_mask).sum())} row(s) where all selected Y = zero/null.")
#         df_f = df_f.loc[nonzero_mask].reset_index(drop=True)

#     if excluded_y_cols:
#         st.info(f"Excluded numeric columns (all zero / null): {', '.join(excluded_y_cols)}")
#     if df_f.empty:
#         st.warning("No rows left after filters."); st.stop()

#     st.subheader("Data Table")
#     st.dataframe(arrow_safe(df_f), use_container_width=True, height=360)

#     # GHG SECTION
#     if any(bool(ghg_data[lbl]) for lbl in GHG_ROW_LABELS) and months_known:
#         st.markdown("---")
#         st.header("📊 GHG Emission Analysis")

#         display_month_list = [m for m in FISCAL_YEAR_ORDER if m in selected_months]
#         st.caption(f"📅 Showing: **{', '.join(display_month_list)}**")

#         # Metric Cards
#         gei_label = "Emission per ton of Equivalent product"
#         months_to_card = [m for m in display_month_list if m in months_known]
#         if months_to_card:
#             cols_cards = st.columns(min(len(months_to_card), 6))
#             for i, m in enumerate(months_to_card):
#                 v = ghg_data[gei_label].get(m)
#                 delta_val = round(v - baseline_value, 4) if v is not None else None
#                 cols_cards[i % len(cols_cards)].metric(
#                     label=m, value=f"{v:.4f}" if v else "N/A",
#                     delta=f"{delta_val:+.4f} vs baseline" if delta_val is not None else None,
#                     delta_color="inverse"
#                 )

#         # GEI Chart
#         st.markdown("---")
#         st.subheader("🎯 Emission Intensity per Ton — GEI")
#         st.caption(f"🟢 Green solid line = T ({t_line_value:.6f}) | 🔴 Red dashed line = Baseline Target ({baseline_value:.4f}) | Shaded band = zone between T and Target")

#         if ghg_data.get(gei_label):
#             fig_gei = make_gei_chart(
#                 ghg_data=ghg_data,
#                 months_known=months_known,
#                 selected_months=selected_months,
#                 baseline_value=baseline_value,
#                 t_line_value=t_line_value,
#                 add_prediction=show_prediction,
#             )
#             st.plotly_chart(fig_gei, use_container_width=True)

#         # Other GHG Charts (kept exactly same)
#         st.markdown("---")
#         st.subheader("📈 Other GHG Metrics")
#         for cfg in OTHER_CHART_CONFIGS:
#             lbl = cfg["label"]
#             if not ghg_data.get(lbl):
#                 continue
#             available = {m: ghg_data[lbl][m] for m in selected_months if m in ghg_data[lbl]}
#             if not available:
#                 continue
#             fig = make_ghg_chart(ghg_data, lbl, selected_months, ghg_chart_type, cfg["y_axis"], cfg["color"])
#             st.subheader(lbl)
#             st.plotly_chart(fig, use_container_width=True)

#         # Prediction Table
#         if show_prediction:
#             remaining = get_remaining_fiscal_months(months_known)
#             if remaining:
#                 st.markdown("---")
#                 st.subheader("📅 Baseline Prediction — Remaining Fiscal Months")
#                 st.caption(f"Known: **{', '.join(months_known)}** | Predicting: **{', '.join(remaining)}**")
#                 # ... (your original prediction table code remains)
#                 pred_rows = []
#                 for lbl in GHG_ROW_LABELS:
#                     vals = ghg_data[lbl]
#                     if len(vals) < 2: continue
#                     mk = [m for m in FISCAL_YEAR_ORDER if m in vals]
#                     mv = [vals[m] for m in mk]
#                     try:
#                         pred_dict, _, _ = predict_remaining_months(mk, mv, FISCAL_YEAR_ORDER)
#                         for m in remaining:
#                             if m in pred_dict:
#                                 pred_rows.append({"Metric": lbl, "Month": m, "Predicted Value": round(pred_dict[m], 6)})
#                     except:
#                         pass
#                 if pred_rows:
#                     pred_df = pd.DataFrame(pred_rows)
#                     pivot = pred_df.pivot(index="Metric", columns="Month", values="Predicted Value")
#                     st.dataframe(pivot, use_container_width=True)

# else:
#     st.info("⬆️ Upload an Excel file to begin.")

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import re
from math import sin, pi

st.set_page_config(page_title="GHG Summary Analyzer", layout="wide")
st.title("GHG Summary Analyzer")
st.caption("Upload Excel → Full Year Analysis with Wave Prediction")

uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

# ====================== HELPERS ======================
def load_excel(uploaded_file):
    for eng in ("openpyxl", "calamine"):
        try:
            return pd.ExcelFile(uploaded_file, engine=eng), eng
        except Exception:
            continue
    raise ImportError("No Excel engine available.")

def find_sheet(xls):
    if "Summary Sheet" in xls.sheet_names:
        return "Summary Sheet"
    for s in xls.sheet_names:
        if "summary" in s.lower():
            return s
    return xls.sheet_names[0]

def header_detect_clean(df_raw):
    raw = df_raw.dropna(how="all").dropna(axis=1, how="all").copy()
    header_idx = raw.head(min(15, len(raw))).notna().sum(axis=1).idxmax()
    df = raw.copy()
    df.columns = df.loc[header_idx].astype(str).str.strip()
    df = df.loc[header_idx + 1:].reset_index(drop=True)
    seen, cols = {}, []
    for c in df.columns:
        base = c if c and c != "nan" else "Unnamed"
        seen[base] = seen.get(base, 0) + 1
        cols.append(base if seen[base] == 1 else f"{base}_{seen[base]}")
    df.columns = cols
    for c in df.columns:
        parsed = pd.to_numeric(df[c], errors="coerce")
        if parsed.notna().mean() >= 0.6:
            df[c] = parsed
    return df

def prune_columns(df, drop_unnamed=True, null_threshold=0.90, drop_pattern=None):
    drop = []
    for c in df.columns:
        name = str(c).strip()
        if drop_unnamed and (name.lower().startswith("unnamed") or name == "" or name.lower() == "nan"):
            drop.append(c); continue
        if drop_pattern and re.match(drop_pattern, name):
            drop.append(c); continue
        if df[c].isna().mean() >= null_threshold:
            drop.append(c); continue
        if df[c].dtype == "object":
            s = df[c].astype(str).str.strip().replace("nan", "")
            if s.replace("", np.nan).dropna().empty:
                drop.append(c); continue
    return df.drop(columns=drop), drop

def hex_to_rgba(hex_color, alpha):
    try:
        hc = str(hex_color).lstrip("#")
        if len(hc) != 6: return hex_color
        r, g, b = int(hc[0:2], 16), int(hc[2:4], 16), int(hc[4:6], 16)
        return f"rgba({r},{g},{b},{alpha})"
    except:
        return hex_color

# ====================== CONFIG ======================
GHG_ROW_LABELS = {
    "Emission per ton of Equivalent product": "GEI",
    "Total Direct and Indirect Emission": "Total",
    "Total Direct Emission (Scope 1)": "Scope1",
    "Total Indirect Emission (Scope 2)": "Scope2",
    "Total Equivalent Product for GHG Emission Intensity": "Product",
}

FISCAL_YEAR_ORDER = ["April","May","June","July","August","September","October","November","December","January","February","March"]

OTHER_CHART_CONFIGS = [
    {"label": "Total Direct and Indirect Emission", "y_axis": "tCO₂ eq", "color": "#E53935"},
    {"label": "Total Direct Emission (Scope 1)", "y_axis": "tCO₂ eq", "color": "#FF7043"},
    {"label": "Total Indirect Emission (Scope 2)", "y_axis": "tCO₂ eq", "color": "#7B1FA2"},
    {"label": "Total Equivalent Product for GHG Emission Intensity", "y_axis": "Tonnes", "color": "#2E7D32"},
]

# ====================== PARSER ======================
def extract_ghg_data(raw):
    month_cols = {}
    for i, row in raw.iterrows():
        for col_idx, val in enumerate(row):
            vs = str(val).strip()
            for m in ["January","February","March","April","May","June","July","August","September","October","November","December"]:
                if m.lower() == vs.lower():
                    month_cols[m] = col_idx
        if len(month_cols) >= 2:
            break

    results = {label: {} for label in GHG_ROW_LABELS}
    for i, row in raw.iterrows():
        row_vals = [str(v).strip() for v in row]
        for label in GHG_ROW_LABELS:
            if any(label.lower() in rv.lower() for rv in row_vals):
                for month, col_idx in month_cols.items():
                    try:
                        v = pd.to_numeric(row.iloc[col_idx], errors="coerce")
                        if pd.notna(v):
                            results[label][month] = v
                    except:
                        pass
    return results, [m for m in FISCAL_YEAR_ORDER if m in month_cols]

# ====================== WAVE PREDICTION ======================
def predict_wave_from_3months(ghg_data, label, months_known, amplitude=0.018):
    if not months_known:
        return {}
    known_vals = [ghg_data[label].get(m) for m in months_known if ghg_data[label].get(m) is not None]
    if not known_vals:
        return {}
    
    avg_value = np.mean(known_vals)
    remaining = [m for m in FISCAL_YEAR_ORDER if m not in months_known]
    
    pred = {}
    for i, m in enumerate(remaining):
        wave = amplitude * sin(2 * pi * i / max(1, len(remaining)))
        pred[m] = round(float(avg_value + wave), 6)
    return pred

# ====================== CHARTS ======================
def make_full_gei_chart(ghg_data, months_known, baseline_value, t_line_value, show_prediction=True):
    label = "Emission per ton of Equivalent product"
    display_months = FISCAL_YEAR_ORDER

    actual_x = [m for m in months_known if m in ghg_data.get(label, {})]
    actual_y = [ghg_data[label][m] for m in actual_x]

    pred_dict = predict_wave_from_3months(ghg_data, label, months_known, amplitude=0.015) if show_prediction else {}
    pred_x = list(pred_dict.keys())
    pred_y = list(pred_dict.values())

    all_vals = actual_y + pred_y + [t_line_value, baseline_value]
    y_range = [min(all_vals or [0])-0.02, max(all_vals or [1])+0.02]

    fig = go.Figure()

    # Shaded Zone
    fig.add_trace(go.Scatter(x=display_months, y=[t_line_value]*12, mode="lines", line=dict(width=0), showlegend=False))
    fig.add_trace(go.Scatter(x=display_months, y=[baseline_value]*12, mode="lines", line=dict(width=0),
                             fill="tonexty", fillcolor="rgba(100,220,130,0.25)", name="Zone: T ↔ Target"))

    # Actual GEI
    if actual_x:
        fig.add_trace(go.Scatter(
            x=actual_x, y=actual_y, name="Actual GEI",
            mode="lines+markers+text", line=dict(color="#00E676", width=4, shape="spline"),
            marker=dict(size=10, color="#00E676", line=dict(color="white", width=2)),
            text=[f"{v:.4f}" for v in actual_y], textposition="top center"
        ))

    # Wave Prediction with smooth join
    if pred_x and show_prediction:
        if actual_x:
            fig.add_trace(go.Scatter(
                x=[actual_x[-1], pred_x[0]], y=[actual_y[-1], pred_y[0]],
                mode="lines", line=dict(color="#00E676", width=3, dash="dot"), showlegend=False
            ))
        fig.add_trace(go.Scatter(
            x=pred_x, y=pred_y, name="Predicted GEI (Wave)",
            mode="lines+markers+text", line=dict(color="#00E676", width=3, dash="dot", shape="spline"),
            marker=dict(size=9, color="#00E676", symbol="diamond"),
            text=[f"{v:.4f}" for v in pred_y], textposition="top center"
        ))

    # T Line
    fig.add_trace(go.Scatter(x=display_months, y=[t_line_value]*12,
        name=f"T = {t_line_value:.6f}", mode="lines+text",
        line=dict(color="#43A047", width=3),
        text=[""]*11 + [f"T = {t_line_value:.4f}"], textposition="middle right"))

    # Baseline Target (Static 0.3552)
    fig.add_trace(go.Scatter(x=display_months, y=[baseline_value]*12,
        name=f"Baseline Target = {baseline_value:.4f}", mode="lines+text",
        line=dict(color="#F44336", width=3, dash="dash"),
        text=[""]*11 + [f"Target = {baseline_value:.4f}"], textposition="middle right"))

    fig.update_layout(
        title="GHG Emission Intensity — Actual GEI · T Line (Green) · Baseline Target (Red)",
        xaxis=dict(title="Month (Fiscal Year Apr → Mar)", categoryorder="array", categoryarray=display_months),
        yaxis=dict(title="GEI (tCO₂ eq/ton)", range=y_range),
        height=580,
        hovermode="x unified",
        legend=dict(orientation="h", y=1.02, x=1, xanchor="right"),
        plot_bgcolor="#0f172a",
        paper_bgcolor="#020617",
        font=dict(color="#e2e8f0")
    )
    return fig

# Other Charts - Respect Chart Type + Wave Prediction + Joined Lines
def make_full_other_chart(ghg_data, label, chart_type, y_axis_label, color, show_prediction=True):
    display_months = FISCAL_YEAR_ORDER
    actual_y = [ghg_data.get(label, {}).get(m) for m in display_months]

    fig = go.Figure()

    # Actual Data
    if chart_type == "Bar":
        fig.add_trace(go.Bar(
            x=display_months, 
            y=[v if v is not None else 0 for v in actual_y],
            name="Actual", marker_color=color
        ))
    elif chart_type == "Area":
        fig.add_trace(go.Scatter(
            x=display_months, y=actual_y, name="Actual",
            mode="lines+markers", fill="tozeroy",
            line=dict(color=color, width=3, shape="spline")
        ))
    else:  # Line
        fig.add_trace(go.Scatter(
            x=display_months, y=actual_y, name="Actual",
            mode="lines+markers", line=dict(color=color, width=3, shape="spline")
        ))

    # Wave Prediction (joined)
    if show_prediction:
        pred_dict = predict_wave_from_3months(ghg_data, label, ["April","May","June"], amplitude=0.08)
        pred_x = list(pred_dict.keys())
        pred_y = list(pred_dict.values())
        if pred_x:
            last_actual_idx = len([v for v in actual_y if v is not None]) - 1
            if last_actual_idx >= 0 and actual_y[last_actual_idx] is not None:
                fig.add_trace(go.Scatter(
                    x=[display_months[last_actual_idx], pred_x[0]],
                    y=[actual_y[last_actual_idx], pred_y[0]],
                    mode="lines", line=dict(color=color, width=3, dash="dot"), showlegend=False
                ))
            fig.add_trace(go.Scatter(
                x=pred_x, y=pred_y, name="Predicted (Wave)",
                mode="lines+markers", line=dict(color=color, width=3, dash="dot", shape="spline"),
                marker=dict(size=8, symbol="diamond")
            ))

    fig.update_layout(
        title=f"{label} — Full Fiscal Year with Wave Prediction",
        xaxis=dict(title="Month (Fiscal Year Apr → Mar)", categoryorder="array", categoryarray=display_months),
        yaxis=dict(title=y_axis_label),
        height=420,
        hovermode="x unified",
        plot_bgcolor="#0f172a",
        paper_bgcolor="#020617"
    )
    return fig

# ====================== MAIN APP ======================
if uploaded:
    try:
        xls, engine_used = load_excel(uploaded)
        sheet = find_sheet(xls)
        raw = pd.read_excel(uploaded, sheet_name=sheet, engine=engine_used, header=None)
        df = header_detect_clean(raw)
        df, _ = prune_columns(df)

        ghg_data, months_known = extract_ghg_data(raw)

        st.success(f"✅ Loaded: **{sheet}** | Detected months: **{months_known}**")

        with st.sidebar:
            st.header("Controls")
            t_line_value = st.number_input("T Line (Green)", value=0.315300, format="%.6f", step=0.0001)
            show_prediction = st.checkbox("Show Wave Prediction", value=True)
            ghg_chart_type = st.radio("Other Charts Type", ["Line", "Bar", "Area"], index=2)

            st.markdown("---")
            st.subheader("📅 Month Filter")
            selected_months = st.multiselect(
                "Select Months to Display",
                options=FISCAL_YEAR_ORDER,
                default=FISCAL_YEAR_ORDER,
                placeholder="Choose months..."
            )

        baseline_value = 0.3552

        # GEI Chart
        st.markdown("---")
        st.subheader("🎯 Emission Intensity per Ton — GEI")
        fig_gei = make_full_gei_chart(ghg_data, months_known, baseline_value, t_line_value, show_prediction)
        st.plotly_chart(fig_gei, use_container_width=True)

        # Prediction Table - Show ALL metrics based on selection
        if show_prediction and months_known:
            st.markdown("---")
            st.subheader("📊 Wave Prediction Table (All Metrics)")
            st.caption(f"Based on average of **{', '.join(months_known)}** + Wave")

            table_data = []
            for lbl_name in GHG_ROW_LABELS:
                if lbl_name not in ghg_data:
                    continue
                pred_dict = predict_wave_from_3months(ghg_data, lbl_name, months_known)
                for month, val in pred_dict.items():
                    table_data.append({
                        "Metric": lbl_name,
                        "Month": month,
                        "Predicted Value": val
                    })

            if table_data:
                pred_df = pd.DataFrame(table_data)
                pivot = pred_df.pivot(index="Metric", columns="Month", values="Predicted Value")
                st.dataframe(pivot.style.format("{:.4f}"), use_container_width=True)

        # Other GHG Charts
        st.markdown("---")
        st.subheader("📈 Other GHG Metrics — Full Year with Wave Prediction")
        for cfg in OTHER_CHART_CONFIGS:
            lbl = cfg["label"]
            if lbl in ghg_data and ghg_data[lbl]:
                st.subheader(lbl)
                fig = make_full_other_chart(
                    ghg_data, lbl, ghg_chart_type, cfg["y_axis"], cfg["color"], show_prediction
                )
                st.plotly_chart(fig, use_container_width=True)

    except Exception as e:
        st.error(f"Error: {e}")

else:
    st.info("⬆️ Upload your Excel file to begin analysis.")