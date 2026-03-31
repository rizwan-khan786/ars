
# # import streamlit as st
# # import pandas as pd
# # import numpy as np
# # import plotly.express as px
# # import re

# # st.set_page_config(page_title="GHG Summary Analyzer", layout="wide")
# # st.title("GHG Summary Analyzer")
# # st.caption("Upload an Excel; we’ll detect the Summary sheet, clean it, and build interactive charts.")

# # uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

# # # ---------------- Helpers ----------------
# # def load_excel(uploaded_file):
# #     for eng in ("openpyxl", "calamine"):
# #         try:
# #             return pd.ExcelFile(uploaded_file, engine=eng), eng
# #         except Exception:
# #             continue
# #     raise ImportError(
# #         "No Excel engine available. Install one of:\n"
# #         "  pip3 install openpyxl\n"
# #         "  or\n"
# #         "  pip3 install pandas-calamine"
# #     )

# # def find_sheet(xls):
# #     if "Summary Sheet" in xls.sheet_names:
# #         return "Summary Sheet"
# #     for s in xls.sheet_names:
# #         if "summary" in s.lower():
# #             return s
# #     return xls.sheet_names[0]

# # def header_detect_clean(df_raw: pd.DataFrame) -> pd.DataFrame:
# #     raw = df_raw.dropna(how="all").dropna(axis=1, how="all").copy()
# #     header_idx = raw.head(min(15, len(raw))).notna().sum(axis=1).idxmax()
# #     df = raw.copy()
# #     df.columns = df.loc[header_idx].astype(str).str.strip()
# #     df = df.loc[header_idx + 1 :].reset_index(drop=True)

# #     # make unique column names
# #     seen, cols = {}, []
# #     for c in df.columns:
# #         base = c if c and c != "nan" else "Unnamed"
# #         seen[base] = seen.get(base, 0) + 1
# #         cols.append(base if seen[base] == 1 else f"{base}_{seen[base]}")
# #     df.columns = cols

# #     # numeric coercion when >=60% convertible
# #     for c in df.columns:
# #         parsed = pd.to_numeric(df[c], errors="coerce")
# #         if parsed.notna().mean() >= 0.6:
# #             df[c] = parsed
# #     return df

# # def prune_columns(df: pd.DataFrame, drop_unnamed=True, null_threshold=0.90, drop_pattern=None):
# #     drop = []
# #     for c in df.columns:
# #         name = str(c).strip()
# #         if drop_unnamed and (name.lower().startswith("unnamed") or name == "" or name.lower() == "nan"):
# #             drop.append(c)
# #             continue
# #         if drop_pattern and re.match(drop_pattern, name):
# #             drop.append(c)
# #             continue
# #         if df[c].isna().mean() >= null_threshold:
# #             drop.append(c)
# #             continue
# #         if df[c].dtype == "object":
# #             s = df[c].astype(str).str.strip().replace("nan", "")
# #             if s.replace("", np.nan).dropna().empty:
# #                 drop.append(c)
# #                 continue
# #     df_clean = df.drop(columns=drop)
# #     return df_clean, drop

# # def guess_label_cols(df, num_cols):
# #     return [c for c in df.columns if c not in num_cols]

# # def arrow_safe(df: pd.DataFrame) -> pd.DataFrame:
# #     out = df.copy()
# #     for c in out.columns:
# #         if out[c].dtype == "object":
# #             out[c] = out[c].astype("string")
# #     return out

# # def reduce_ticks(labels, max_ticks=25):
# #     n = len(labels)
# #     if n <= max_ticks:
# #         return list(range(n)), labels
# #     step = max(1, n // max_ticks)
# #     idxs = list(range(0, n, step))
# #     return idxs, [labels[i] for i in idxs]

# # def hex_to_rgba(hex_color: str, alpha: float) -> str:
# #     """Convert #RRGGBB to rgba(...) string (alpha 0..1).
# #        If the color is not a hex string, return the original color (Plotly will use it as-is)."""
# #     try:
# #         hc = str(hex_color).lstrip("#")
# #         if len(hc) != 6:
# #             return hex_color
# #         r = int(hc[0:2], 16)
# #         g = int(hc[2:4], 16)
# #         b = int(hc[4:6], 16)
# #         return f"rgba({r},{g},{b},{alpha})"
# #     except Exception:
# #         return hex_color

# # def palette_color(palette: list, idx: int) -> str:
# #     return palette[idx % len(palette)]

# # PALETTES = {
# #     "Plotly": px.colors.qualitative.Plotly,
# #     "D3": px.colors.qualitative.D3,
# #     "Bold": px.colors.qualitative.Bold,
# #     "Dark24": px.colors.qualitative.Dark24,
# #     "G10": px.colors.qualitative.G10,
# #     "Alphabet": px.colors.qualitative.Alphabet,
# # }

# # # ---------------- App ----------------
# # if uploaded:
# #     try:
# #         xls, engine_used = load_excel(uploaded)
# #         st.caption(f"Excel engine: **{engine_used}**")
# #         sheet = find_sheet(xls)
# #         st.success(f"Using sheet: {sheet}")
# #         raw = pd.read_excel(uploaded, sheet_name=sheet, engine=engine_used, header=None)
# #         df = header_detect_clean(raw)
# #     except Exception as e:
# #         st.exception(e)
# #         st.stop()

# #     # prune junk columns (Unnamed, mostly-empty, whitespace-only, 0.0_* pattern)
# #     drop_pattern = r"^0\.0_.*"
# #     df, dropped_prune = prune_columns(df, drop_unnamed=True, null_threshold=0.90, drop_pattern=drop_pattern)
# #     if dropped_prune:
# #         st.info(f"Hidden {len(dropped_prune)} mostly-empty / unnamed / patterned columns.")

# #     # detect numeric columns after pruning
# #     num_cols = df.select_dtypes(include="number").columns.tolist()

# #     # filter numeric columns to exclude any that are all-NaN OR all-zero
# #     valid_y_cols = []
# #     excluded_y_cols = []
# #     for c in num_cols:
# #         s = pd.to_numeric(df[c], errors="coerce")
# #         s_non_na = s.dropna()
# #         if s_non_na.empty:
# #             excluded_y_cols.append(c)
# #         elif s_non_na.abs().sum() == 0:
# #             excluded_y_cols.append(c)
# #         else:
# #             valid_y_cols.append(c)

# #     # label/categorical candidates
# #     cat_cols = guess_label_cols(df, num_cols)

# #     # ---------------- SIDEBAR ----------------
# #     with st.sidebar:
# #         st.header("Controls")

# #         # X-axis chooser
# #         x_col = st.selectbox("X-axis (categorical / time)", options=(cat_cols or [None]))

# #         # Y-axis chooser: only show numeric columns that have at least one non-zero value
# #         if valid_y_cols:
# #             y_cols = st.multiselect("Y-axis (numeric)", options=valid_y_cols, default=valid_y_cols[:1])
# #         else:
# #             st.warning("No numeric columns with non-zero/nonnull values were detected.")
# #             y_cols = []

# #         agg = st.radio("Aggregate by X", ["None", "sum", "mean"], index=0)
# #         top_n = st.number_input("Limit to Top N (by first Y)", min_value=0, value=0, step=1, help="0 = show all")

# #         # Palette & custom color controls
# #         st.markdown("### Colors")
# #         palette_name = st.selectbox("Palette", options=list(PALETTES.keys()), index=0)
# #         palette = PALETTES[palette_name]
# #         use_custom_colors = st.checkbox("Use custom colors", value=False)

# #         # dynamic per-Y color pickers (appear only if custom colors enabled)
# #         custom_color_map = {}
# #         if use_custom_colors and y_cols:
# #             st.markdown("Pick a color for each selected Y:")
# #             for i, y in enumerate(y_cols):
# #                 default_color = palette_color(palette, i)
# #                 color_val = st.color_picker(f"{y}", value=default_color, key=f"col_picker_{y}")
# #                 custom_color_map[y] = color_val

# #         # build filter options after removing zero-only rows for selected Y columns
# #         if y_cols:
# #             y_numeric_all = df[y_cols].apply(pd.to_numeric, errors="coerce").fillna(0).abs()
# #             nonzero_mask_for_filters = (y_numeric_all.sum(axis=1) != 0)
# #             df_for_filters = df.loc[nonzero_mask_for_filters].copy()
# #         else:
# #             df_for_filters = df.copy()

# #         st.markdown("**Filters**")
# #         active_filters = {}
# #         for c in cat_cols[:6]:
# #             vals = df_for_filters[c].dropna().astype(str).unique().tolist()
# #             if 1 < len(vals) <= 200:
# #                 sel = st.multiselect(f"{c}", options=vals, default=[], placeholder="(all)")
# #                 if sel:
# #                     active_filters[c] = set(sel)

# #         chart_type = st.radio("Chart type", ["Line", "Bar", "Area"], index=1)
# #         use_log = st.checkbox("Log scale (Y)", value=False)
# #         rolling = st.number_input("Rolling mean (window)", min_value=0, value=0, step=1)
# #         max_ticks = st.slider("Max X ticks", min_value=5, max_value=60, value=25)

# #     # ---------------- APPLY FILTERS & AGGREGATION ----------------
# #     filt = pd.Series(True, index=df.index)
# #     for col, allowed in active_filters.items():
# #         filt &= df[col].astype(str).isin(allowed)
# #     df_f = df.loc[filt].copy()

# #     if x_col and agg != "None" and y_cols:
# #         if agg == "sum":
# #             df_f = df_f.groupby(x_col, dropna=False)[y_cols].sum().reset_index()
# #         else:
# #             df_f = df_f.groupby(x_col, dropna=False)[y_cols].mean().reset_index()

# #     if top_n and y_cols and x_col and x_col in df_f.columns:
# #         df_f = df_f.sort_values(y_cols[0], ascending=False).head(int(top_n))

# #     # remove rows where ALL selected Y columns are zero or null
# #     if y_cols:
# #         y_numeric_after = df_f[y_cols].apply(pd.to_numeric, errors="coerce")
# #         nonzero_mask_after = (y_numeric_after.fillna(0).abs().sum(axis=1) != 0)
# #         removed_count = (~nonzero_mask_after).sum()
# #         if removed_count:
# #             st.info(f"Removed {int(removed_count)} row(s) where selected Y columns were all zero/null.")
# #         df_f = df_f.loc[nonzero_mask_after].reset_index(drop=True)

# #     if excluded_y_cols:
# #         st.info(f"Excluded numeric columns (all zero or all null): {', '.join(excluded_y_cols)}")

# #     if df_f.empty:
# #         st.warning("No rows left after applying filters / removing zero rows. Adjust filters or Y selection.")
# #         st.stop()

# #     # Table
# #     st.subheader("Table")
# #     st.dataframe(arrow_safe(df_f), width="stretch", height=420)

# #     if not y_cols:
# #         st.warning("Pick at least one numeric Y column.")
# #         st.stop()

# #     # Build x values
# #     if x_col and x_col in df_f.columns:
# #         x_vals = df_f[x_col].astype(str).fillna("")
# #     else:
# #         x_vals = pd.Series(range(len(df_f)), index=df_f.index, name="Index")
# #         x_col = "Index"
# #         df_f = df_f.assign(Index=x_vals)

# #     # Plotting: apply colors with correct trace properties per chart type
# #     for i, y in enumerate(y_cols):
# #         y_vals = pd.to_numeric(df_f[y], errors="coerce")
# #         if y_vals.dropna().empty or y_vals.dropna().abs().sum() == 0:
# #             st.info(f"Skipping chart for '{y}' (all values are zero or null after filters).")
# #             continue

# #         if use_custom_colors and y in custom_color_map:
# #             primary_color = custom_color_map[y]
# #         else:
# #             primary_color = palette_color(palette, i)

# #         y_vals_filled = y_vals.fillna(0)
# #         df_plot = pd.DataFrame({x_col: x_vals, y: y_vals_filled})

# #         if chart_type == "Bar":
# #             fig = px.bar(df_plot, x=x_col, y=y)
# #             # bar traces use marker.color or marker=dict(color=...)
# #             fig.update_traces(marker=dict(color=primary_color))
# #         elif chart_type == "Line":
# #             fig = px.line(df_plot, x=x_col, y=y)
# #             # line traces accept line.color
# #             fig.update_traces(line=dict(color=primary_color))
# #         else:  # Area
# #             fig = px.area(df_plot, x=x_col, y=y)
# #             # area is a scatter with fill; set line color and a semi-transparent fillcolor
# #             fig.update_traces(line=dict(color=primary_color),
# #                               fillcolor=hex_to_rgba(primary_color, 0.25))

# #         # Rolling mean overlay — dashed semi-transparent line
# #         if rolling and rolling > 1:
# #             roll = y_vals_filled.rolling(int(rolling), min_periods=1).mean()
# #             roll_color = hex_to_rgba(primary_color, 0.8)
# #             fig.add_scatter(x=df_plot[x_col], y=roll, mode="lines",
# #                             name=f"{y} (rolling {int(rolling)})",
# #                             line=dict(color=roll_color, dash="dash"))

# #         idxs, labels = reduce_ticks(df_plot[x_col].astype(str).tolist(), max_ticks=max_ticks)

# #         fig.update_layout(
# #             xaxis=dict(tickmode="array", tickvals=[df_plot[x_col].iloc[i] for i in idxs], ticktext=labels),
# #             yaxis=dict(type="log" if use_log else "linear"),
# #             margin=dict(l=10, r=10, t=40, b=10),
# #             height=480,
# #             hovermode="x unified",
# #         )

# #         st.plotly_chart(fig, use_container_width=True)

# #     # Download cleaned/filtered data
# #     st.download_button(
# #         "Download cleaned/filtered data (CSV)",
# #         df_f.to_csv(index=False).encode("utf-8"),
# #         file_name="summary_cleaned_filtered.csv",
# #         mime="text/csv",
# #     )

# # else:
# #     st.info("Upload a file to begin.")


# import streamlit as st
# import pandas as pd
# import numpy as np
# import plotly.express as px
# import re
# from datetime import datetime

# st.set_page_config(page_title="GHG Summary Analyzer", layout="wide")
# st.title("GHG Summary Analyzer")
# st.caption("Upload one or more Excel files • Auto-detects Summary sheet • Interactive charts + Month-wise view")

# # ---------------- Helpers (unchanged) ----------------
# def load_excel(uploaded_file):
#     for eng in ("openpyxl", "calamine"):
#         try:
#             return pd.ExcelFile(uploaded_file, engine=eng), eng
#         except Exception:
#             continue
#     raise ImportError("No Excel engine available. Install openpyxl or pandas-calamine")

# def find_sheet(xls):
#     if "Summary Sheet" in xls.sheet_names:
#         return "Summary Sheet"
#     for s in xls.sheet_names:
#         if "summary" in s.lower():
#             return s
#     return xls.sheet_names[0]

# def header_detect_clean(df_raw: pd.DataFrame) -> pd.DataFrame:
#     raw = df_raw.dropna(how="all").dropna(axis=1, how="all").copy()
#     header_idx = raw.head(min(15, len(raw))).notna().sum(axis=1).idxmax()
#     df = raw.copy()
#     df.columns = df.loc[header_idx].astype(str).str.strip()
#     df = df.loc[header_idx + 1 :].reset_index(drop=True)

#     # Make unique column names
#     seen, cols = {}, []
#     for c in df.columns:
#         base = c if c and c != "nan" else "Unnamed"
#         seen[base] = seen.get(base, 0) + 1
#         cols.append(base if seen[base] == 1 else f"{base}_{seen[base]}")
#     df.columns = cols

#     # Numeric coercion when >=60% convertible
#     for c in df.columns:
#         parsed = pd.to_numeric(df[c], errors="coerce")
#         if parsed.notna().mean() >= 0.6:
#             df[c] = parsed
#     return df

# def prune_columns(df: pd.DataFrame, drop_unnamed=True, null_threshold=0.90, drop_pattern=None):
#     drop = []
#     for c in df.columns:
#         name = str(c).strip()
#         if drop_unnamed and (name.lower().startswith("unnamed") or name == "" or name.lower() == "nan"):
#             drop.append(c)
#             continue
#         if drop_pattern and re.match(drop_pattern, name):
#             drop.append(c)
#             continue
#         if df[c].isna().mean() >= null_threshold:
#             drop.append(c)
#             continue
#         if df[c].dtype == "object":
#             s = df[c].astype(str).str.strip().replace("nan", "")
#             if s.replace("", np.nan).dropna().empty:
#                 drop.append(c)
#                 continue
#     return df.drop(columns=drop), drop

# def guess_label_cols(df, num_cols):
#     return [c for c in df.columns if c not in num_cols]

# def arrow_safe(df: pd.DataFrame) -> pd.DataFrame:
#     out = df.copy()
#     for c in out.columns:
#         if out[c].dtype == "object":
#             out[c] = out[c].astype("string")
#     return out

# def reduce_ticks(labels, max_ticks=25):
#     n = len(labels)
#     if n <= max_ticks:
#         return list(range(n)), labels
#     step = max(1, n // max_ticks)
#     idxs = list(range(0, n, step))
#     return idxs, [labels[i] for i in idxs]

# def hex_to_rgba(hex_color: str, alpha: float) -> str:
#     try:
#         hc = str(hex_color).lstrip("#")
#         if len(hc) != 6:
#             return hex_color
#         r = int(hc[0:2], 16)
#         g = int(hc[2:4], 16)
#         b = int(hc[4:6], 16)
#         return f"rgba({r},{g},{b},{alpha})"
#     except Exception:
#         return hex_color

# def palette_color(palette: list, idx: int) -> str:
#     return palette[idx % len(palette)]

# PALETTES = {
#     "Plotly": px.colors.qualitative.Plotly,
#     "D3": px.colors.qualitative.D3,
#     "Bold": px.colors.qualitative.Bold,
#     "Dark24": px.colors.qualitative.Dark24,
#     "G10": px.colors.qualitative.G10,
#     "Alphabet": px.colors.qualitative.Alphabet,
# }

# # ---------------- Main App ----------------
# uploaded_files = st.file_uploader(
#     "Upload Excel file(s) (.xlsx)", 
#     type=["xlsx"], 
#     accept_multiple_files=True
# )

# if uploaded_files:
#     all_dfs = []

#     for uploaded in uploaded_files:
#         try:
#             xls, engine_used = load_excel(uploaded)
#             sheet = find_sheet(xls)
#             raw = pd.read_excel(uploaded, sheet_name=sheet, engine=engine_used, header=None)
#             df = header_detect_clean(raw)
            
#             # Prune junk columns
#             drop_pattern = r"^0\.0_.*"
#             df, _ = prune_columns(df, drop_unnamed=True, null_threshold=0.90, drop_pattern=drop_pattern)
            
#             df['Source_File'] = uploaded.name  # Track which file the data came from
#             all_dfs.append(df)
            
#         except Exception as e:
#             st.error(f"Error processing {uploaded.name}: {e}")

#     if not all_dfs:
#         st.stop()

#     # Combine all files
#     df_combined = pd.concat(all_dfs, ignore_index=True)

#     # Detect numeric columns
#     num_cols = df_combined.select_dtypes(include="number").columns.tolist()

#     # Filter valid Y columns (not all zero/null)
#     valid_y_cols = []
#     for c in num_cols:
#         s = pd.to_numeric(df_combined[c], errors="coerce")
#         s_non_na = s.dropna()
#         if not s_non_na.empty and s_non_na.abs().sum() != 0:
#             valid_y_cols.append(c)

#     cat_cols = [c for c in df_combined.columns if c not in num_cols and c != 'Source_File']

#     # ---------------- SIDEBAR ----------------
#     with st.sidebar:
#         st.header("Controls")

#         x_col = st.selectbox("X-axis (categorical / time)", options=cat_cols or [None])

#         if valid_y_cols:
#             y_cols = st.multiselect("Y-axis (numeric)", options=valid_y_cols, default=valid_y_cols[:1])
#         else:
#             st.warning("No valid numeric columns found.")
#             y_cols = []

#         agg = st.radio("Aggregate by X", ["None", "sum", "mean"], index=0)
#         top_n = st.number_input("Limit to Top N", min_value=0, value=0, step=1)

#         st.markdown("### Colors")
#         palette_name = st.selectbox("Palette", options=list(PALETTES.keys()), index=0)
#         palette = PALETTES[palette_name]
#         use_custom_colors = st.checkbox("Use custom colors", value=False)

#         custom_color_map = {}
#         if use_custom_colors and y_cols:
#             for i, y in enumerate(y_cols):
#                 default_color = palette_color(palette, i)
#                 color_val = st.color_picker(f"Color for {y}", value=default_color, key=f"col_{y}")
#                 custom_color_map[y] = color_val

#         # Filters
#         st.markdown("**Filters**")
#         active_filters = {}
#         df_for_filters = df_combined.copy()
#         for c in cat_cols[:6]:
#             vals = df_for_filters[c].dropna().astype(str).unique().tolist()
#             if 1 < len(vals) <= 200:
#                 sel = st.multiselect(f"Filter {c}", options=vals, default=[], placeholder="(all)")
#                 if sel:
#                     active_filters[c] = set(sel)

#         chart_type = st.radio("Chart type", ["Line", "Bar", "Area"], index=1)
#         use_log = st.checkbox("Log scale (Y)", value=False)
#         rolling = st.number_input("Rolling mean window", min_value=0, value=0, step=1)
#         max_ticks = st.slider("Max X ticks", 5, 60, 25)

#     # ---------------- Apply Filters & Aggregation ----------------
#     filt = pd.Series(True, index=df_combined.index)
#     for col, allowed in active_filters.items():
#         filt &= df_combined[col].astype(str).isin(allowed)
#     df_f = df_combined.loc[filt].copy()

#     if x_col and agg != "None" and y_cols:
#         if agg == "sum":
#             df_f = df_f.groupby(x_col, dropna=False)[y_cols].sum().reset_index()
#         else:
#             df_f = df_f.groupby(x_col, dropna=False)[y_cols].mean().reset_index()

#     if top_n and y_cols and x_col:
#         df_f = df_f.sort_values(y_cols[0], ascending=False).head(int(top_n))

#     # Remove all-zero rows for selected Y columns
#     if y_cols:
#         y_num = df_f[y_cols].apply(pd.to_numeric, errors="coerce").fillna(0)
#         df_f = df_f[y_num.abs().sum(axis=1) != 0].reset_index(drop=True)

#     if df_f.empty:
#         st.warning("No data left after filtering.")
#         st.stop()

#     # Table
#     st.subheader("Cleaned & Filtered Data")
#     st.dataframe(arrow_safe(df_f), use_container_width=True, height=400)

#     if not y_cols:
#         st.warning("Please select at least one Y column.")
#         st.stop()

#     # ==================== ORIGINAL CHARTS ====================
#     st.subheader("Main Charts")
#     x_vals = df_f[x_col].astype(str).fillna("") if x_col else pd.Series(range(len(df_f)))

#     for i, y in enumerate(y_cols):
#         y_vals = pd.to_numeric(df_f[y], errors="coerce").fillna(0)

#         if y_vals.abs().sum() == 0:
#             continue

#         color = custom_color_map.get(y) if use_custom_colors else palette_color(palette, i)

#         df_plot = pd.DataFrame({x_col or "Index": x_vals, y: y_vals})

#         if chart_type == "Bar":
#             fig = px.bar(df_plot, x=df_plot.columns[0], y=y)
#             fig.update_traces(marker=dict(color=color))
#         elif chart_type == "Line":
#             fig = px.line(df_plot, x=df_plot.columns[0], y=y)
#             fig.update_traces(line=dict(color=color))
#         else:  # Area
#             fig = px.area(df_plot, x=df_plot.columns[0], y=y)
#             fig.update_traces(line=dict(color=color), fillcolor=hex_to_rgba(color, 0.25))

#         if rolling > 1:
#             roll = y_vals.rolling(int(rolling), min_periods=1).mean()
#             fig.add_scatter(x=df_plot.iloc[:,0], y=roll, mode="lines",
#                             name=f"{y} (rolling)", line=dict(color=hex_to_rgba(color, 0.8), dash="dash"))

#         idxs, labels = reduce_ticks(df_plot.iloc[:,0].astype(str).tolist(), max_ticks)

#         fig.update_layout(
#             xaxis=dict(tickmode="array", tickvals=[df_plot.iloc[i,0] for i in idxs], ticktext=labels),
#             yaxis=dict(type="log" if use_log else "linear"),
#             height=480,
#             hovermode="x unified",
#             title=f"{y} - {chart_type} Chart"
#         )

#         st.plotly_chart(fig, use_container_width=True)

#     # ==================== NEW: MONTH-WISE CHARTS ====================
#     st.subheader("📅 Month-wise Analysis")

#     # Try to auto-detect a date column
#     date_col = None
#     for col in df_f.columns:
#         if df_f[col].dtype == 'object':
#             try:
#                 pd.to_datetime(df_f[col], errors='coerce')
#                 date_col = col
#                 break
#             except:
#                 pass
#         elif pd.api.types.is_datetime64_any_dtype(df_f[col]):
#             date_col = col
#             break

#     if date_col:
#         df_month = df_f.copy()
#         df_month['Month'] = pd.to_datetime(df_month[date_col], errors='coerce').dt.strftime('%Y-%m')
#         df_month = df_month.dropna(subset=['Month'])

#         if not df_month.empty and y_cols:
#             monthly_agg = df_month.groupby('Month')[y_cols].sum().reset_index()

#             st.caption(f"Grouped by Month from column: **{date_col}**")

#             for y in y_cols:
#                 fig_month = px.bar(monthly_agg, x='Month', y=y, title=f"Month-wise {y} (Sum)")
#                 fig_month.update_traces(marker_color=palette_color(palette, y_cols.index(y)))
#                 fig_month.update_layout(height=450)
#                 st.plotly_chart(fig_month, use_container_width=True)
#         else:
#             st.info("Could not create month-wise view. No valid date column detected or no data.")
#     else:
#         st.info("No date-like column detected for month-wise analysis.")

#     # Download
#     st.download_button(
#         "Download cleaned/filtered data (CSV)",
#         df_f.to_csv(index=False).encode("utf-8"),
#         file_name="ghg_summary_cleaned_filtered.csv",
#         mime="text/csv",
#     )

# else:
#     st.info("Upload one or more Excel files to begin analysis.")

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import re
from datetime import datetime

st.set_page_config(page_title="GHG Summary Analyzer", layout="wide")
st.title("GHG Summary Analyzer")
st.caption("Upload one or more Excel files • Improved Month-wise detection")

# ---------------- Helpers (kept same) ----------------
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

def header_detect_clean(df_raw: pd.DataFrame) -> pd.DataFrame:
    raw = df_raw.dropna(how="all").dropna(axis=1, how="all").copy()
    header_idx = raw.head(min(15, len(raw))).notna().sum(axis=1).idxmax()
    df = raw.copy()
    df.columns = df.loc[header_idx].astype(str).str.strip()
    df = df.loc[header_idx + 1 :].reset_index(drop=True)

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

def prune_columns(df: pd.DataFrame, drop_unnamed=True, null_threshold=0.90, drop_pattern=None):
    drop = []
    for c in df.columns:
        name = str(c).strip()
        if drop_unnamed and (name.lower().startswith("unnamed") or name == "" or name.lower() == "nan"):
            drop.append(c)
            continue
        if drop_pattern and re.match(drop_pattern, name):
            drop.append(c)
            continue
        if df[c].isna().mean() >= null_threshold:
            drop.append(c)
            continue
        if df[c].dtype == "object":
            s = df[c].astype(str).str.strip().replace("nan", "")
            if s.replace("", np.nan).dropna().empty:
                drop.append(c)
                continue
    return df.drop(columns=drop), drop

def arrow_safe(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for c in out.columns:
        if out[c].dtype == "object":
            out[c] = out[c].astype("string")
    return out

def reduce_ticks(labels, max_ticks=25):
    n = len(labels)
    if n <= max_ticks:
        return list(range(n)), labels
    step = max(1, n // max_ticks)
    idxs = list(range(0, n, step))
    return idxs, [labels[i] for i in idxs]

def hex_to_rgba(hex_color: str, alpha: float) -> str:
    try:
        hc = str(hex_color).lstrip("#")
        if len(hc) != 6: return hex_color
        r = int(hc[0:2], 16); g = int(hc[2:4], 16); b = int(hc[4:6], 16)
        return f"rgba({r},{g},{b},{alpha})"
    except:
        return hex_color

def palette_color(palette: list, idx: int) -> str:
    return palette[idx % len(palette)]

PALETTES = {
    "Plotly": px.colors.qualitative.Plotly,
    "D3": px.colors.qualitative.D3,
    "Bold": px.colors.qualitative.Bold,
    "Dark24": px.colors.qualitative.Dark24,
    "G10": px.colors.qualitative.G10,
    "Alphabet": px.colors.qualitative.Alphabet,
}

# ---------------- Main App ----------------
uploaded_files = st.file_uploader("Upload Excel file(s) (.xlsx)", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    all_dfs = []
    for uploaded in uploaded_files:
        try:
            xls, engine_used = load_excel(uploaded)
            sheet = find_sheet(xls)
            raw = pd.read_excel(uploaded, sheet_name=sheet, engine=engine_used, header=None)
            df = header_detect_clean(raw)
            df, _ = prune_columns(df, drop_unnamed=True, null_threshold=0.90, drop_pattern=r"^0\.0_.*")
            df['Source_File'] = uploaded.name
            all_dfs.append(df)
        except Exception as e:
            st.error(f"Error with {uploaded.name}: {e}")

    if not all_dfs:
        st.stop()

    df_combined = pd.concat(all_dfs, ignore_index=True)

    num_cols = df_combined.select_dtypes(include="number").columns.tolist()
    valid_y_cols = [c for c in num_cols 
                    if pd.to_numeric(df_combined[c], errors="coerce").dropna().abs().sum() != 0]

    cat_cols = [c for c in df_combined.columns if c not in num_cols and c != 'Source_File']

    # ---------------- SIDEBAR ----------------
    with st.sidebar:
        st.header("Controls")
        x_col = st.selectbox("X-axis", options=cat_cols or [None])
        y_cols = st.multiselect("Y-axis (numeric)", options=valid_y_cols, 
                               default=valid_y_cols[:1] if valid_y_cols else [])

        agg = st.radio("Aggregate by X", ["None", "sum", "mean"], index=0)
        top_n = st.number_input("Top N", min_value=0, value=0, step=1)

        st.markdown("### Colors")
        palette_name = st.selectbox("Palette", list(PALETTES.keys()), index=0)
        palette = PALETTES[palette_name]
        use_custom_colors = st.checkbox("Use custom colors", value=False)

        custom_color_map = {}
        if use_custom_colors and y_cols:
            for i, y in enumerate(y_cols):
                color_val = st.color_picker(f"{y}", value=palette_color(palette, i), key=f"col_{y}")
                custom_color_map[y] = color_val

        st.markdown("**Filters**")
        active_filters = {}
        for c in cat_cols[:6]:
            vals = df_combined[c].dropna().astype(str).unique().tolist()
            if 1 < len(vals) <= 200:
                sel = st.multiselect(f"Filter {c}", options=vals, default=[], placeholder="(all)")
                if sel:
                    active_filters[c] = set(sel)

        chart_type = st.radio("Chart type", ["Line", "Bar", "Area"], index=1)
        use_log = st.checkbox("Log scale (Y)", value=False)
        rolling = st.number_input("Rolling mean window", min_value=0, value=0, step=1)
        max_ticks = st.slider("Max X ticks", 5, 60, 25)

    # Apply filters & aggregation (same logic)
    filt = pd.Series(True, index=df_combined.index)
    for col, allowed in active_filters.items():
        filt &= df_combined[col].astype(str).isin(allowed)
    df_f = df_combined.loc[filt].copy()

    if x_col and agg != "None" and y_cols:
        df_f = df_f.groupby(x_col, dropna=False)[y_cols].agg(agg).reset_index()

    if top_n and y_cols and x_col:
        df_f = df_f.sort_values(y_cols[0], ascending=False).head(int(top_n))

    if y_cols:
        y_num = df_f[y_cols].apply(pd.to_numeric, errors="coerce").fillna(0)
        df_f = df_f[y_num.abs().sum(axis=1) != 0].reset_index(drop=True)

    if df_f.empty:
        st.warning("No data left after filtering.")
        st.stop()

    st.subheader("Cleaned & Filtered Data")
    st.dataframe(arrow_safe(df_f), use_container_width=True, height=400)

    if not y_cols:
        st.stop()

    # ==================== MAIN CHARTS (unchanged) ====================
    st.subheader("Main Charts")
    x_vals = df_f[x_col].astype(str) if x_col else pd.Series(range(len(df_f)), name="Index")

    for i, y in enumerate(y_cols):
        color = custom_color_map.get(y) if use_custom_colors else palette_color(palette, i)
        df_plot = pd.DataFrame({x_col or "Index": x_vals, y: pd.to_numeric(df_f[y], errors="coerce").fillna(0)})

        if chart_type == "Bar":
            fig = px.bar(df_plot, x=df_plot.columns[0], y=y)
            fig.update_traces(marker_color=color)
        elif chart_type == "Line":
            fig = px.line(df_plot, x=df_plot.columns[0], y=y)
            fig.update_traces(line_color=color)
        else:
            fig = px.area(df_plot, x=df_plot.columns[0], y=y)
            fig.update_traces(line_color=color, fillcolor=hex_to_rgba(color, 0.25))

        if rolling > 1:
            roll = df_plot[y].rolling(rolling, min_periods=1).mean()
            fig.add_scatter(x=df_plot.iloc[:,0], y=roll, mode="lines",
                            name=f"Rolling", line=dict(color=hex_to_rgba(color, 0.8), dash="dash"))

        idxs, labels = reduce_ticks(df_plot.iloc[:,0].tolist(), max_ticks)
        fig.update_layout(
            xaxis=dict(tickmode="array", tickvals=[df_plot.iloc[i,0] for i in idxs], ticktext=labels),
            yaxis_type="log" if use_log else "linear",
            height=480,
            title=f"{y} - {chart_type}"
        )
        st.plotly_chart(fig, use_container_width=True)

    # ==================== IMPROVED MONTH-WISE ANALYSIS ====================
    st.subheader("📅 Month-wise Analysis")

    # Improved detection: look for date-like column headers (2023-04-01, 2023-05, etc.)
    month_col = None
    date_pattern = re.compile(r'^\d{4}-\d{1,2}(-\d{1,2})?')

    for col in df_f.columns:
        col_str = str(col).strip()
        if date_pattern.match(col_str) or any(x in col_str.lower() for x in ['month', 'period', 'date']):
            month_col = col
            break

    if month_col:
        df_month = df_f.copy()
        
        # Convert column name or values to clean YYYY-MM format
        try:
            # If the column itself is a date string (like header "2023-04-01")
            month_key = pd.to_datetime(str(month_col), errors='coerce')
            if pd.notna(month_key):
                df_month['Month'] = month_key.strftime('%Y-%m')
            else:
                # Try parsing values inside the column
                df_month['Month'] = pd.to_datetime(df_month[month_col], errors='coerce').dt.strftime('%Y-%m')
        except:
            df_month['Month'] = df_month[month_col].astype(str).str[:7]  # take first 7 chars (YYYY-MM)

        df_month = df_month.dropna(subset=['Month'])
        
        if not df_month.empty and y_cols:
            monthly_sum = df_month.groupby('Month')[y_cols].sum().reset_index()
            monthly_sum = monthly_sum.sort_values('Month')

            st.success(f"✅ Using **{month_col}** for month-wise view")

            for y in y_cols:
                fig_m = px.bar(monthly_sum, x='Month', y=y,
                              title=f"Month-wise {y} (Total)",
                              labels={'Month': 'Month'})
                fig_m.update_traces(marker_color=palette_color(palette, y_cols.index(y)))
                fig_m.update_layout(height=450)
                st.plotly_chart(fig_m, use_container_width=True)
        else:
            st.info("No data available for month-wise after grouping.")
    else:
        st.info("Still could not detect a month/date column. "
                "Please rename the column containing months (e.g. 2023-04-01) to **'Month'** and re-upload.")

    # Download
    st.download_button(
        "Download cleaned/filtered data (CSV)",
        df_f.to_csv(index=False).encode("utf-8"),
        file_name="ghg_summary_cleaned.csv",
        mime="text/csv"
    )

else:
    st.info("Upload your Excel file(s) to begin.")