# app.py
# Streamlit app: Convert "Smith Integrated Care Services - Report Time" → SOTA-style workbook
# Author: you + ChatGPT
#
# How it works
# 1) Upload the Smith Integrated Excel (typically a single sheet called "Worksheet").
# 2) App detects headers (promotes the first row with "Date Of Visit" to header if needed).
# 3) You pick the year (default 2025).
# 4) We compute SOTA weeks by day-of-month: W1=1–7, W2=8–14, W3=15–21, W4=22–28, W5=29–end.
# 5) It outputs a workbook: one sheet per employee. Each sheet is grouped by Month with a banner row,
#    client rows, columns Week 1..5 and Month Total. Optional Totals row and optional "#days worked" row.

import io
import re
import pandas as pd
import numpy as np
import streamlit as st
from collections import OrderedDict

st.set_page_config(page_title="SOTA Converter", page_icon="🗂️", layout="wide")

# ----------------------------
# Helper functions
# ----------------------------
MONTHS = ["January","February","March","April","May","June",
          "July","August","September","October","November","December"]
MONTH_TO_NUM = {m: i+1 for i, m in enumerate(MONTHS)}

def sota_week(day: int) -> int:
    """Map calendar day-of-month to SOTA week 1..5."""
    if pd.isna(day):
        return np.nan
    d = int(day)
    if 1 <= d <= 7: return 1
    if 8 <= d <= 14: return 2
    if 15 <= d <= 21: return 3
    if 22 <= d <= 28: return 4
    if d >= 29: return 5
    return np.nan

def pick_col(df: pd.DataFrame, candidates, fuzzy_contains=False):
    """
    Pick the first matching column from `df` for the set of candidate names.
    If fuzzy_contains=True, match if candidate substring appears in col name (case-insensitive).
    """
    cols = list(df.columns)
    cl = [str(c).strip().lower() for c in cols]
    for cand in candidates:
        lc = cand.lower()
        if fuzzy_contains:
            for i, c in enumerate(cl):
                if lc in c:
                    return cols[i]
        else:
            for i, c in enumerate(cl):
                if lc == c:
                    return cols[i]
    return None

def load_smith_dataframe(file) -> pd.DataFrame:
    """
    Read the source Excel and return a clean DataFrame with columns:
      - 'Date Of Visit' (datetime)
      - 'User' (employee)
      - 'Client' (client)
      - 'Hours' (float)
    The Smith export often has the first row as a banner header row; we promote it.
    """
    xls = pd.ExcelFile(file)
    # Prefer 'Worksheet' if present
    sheet = "Worksheet" if "Worksheet" in xls.sheet_names else xls.sheet_names[0]
    df_raw = pd.read_excel(file, sheet_name=sheet, header=None)

    # Find a header row containing "Date Of Visit"
    header_row_candidates = df_raw.index[df_raw.apply(lambda r: r.astype(str).str.contains("Date Of Visit", na=False)).any(axis=1)].tolist()
    if header_row_candidates:
        hidx = header_row_candidates[0]
        headers = df_raw.iloc[hidx].tolist()
        df = df_raw.iloc[hidx+1:].copy()
        df.columns = headers
    else:
        # Fallback: read with headers normally
        df = pd.read_excel(file, sheet_name=sheet)

    # Normalize key columns (robust picking)
    date_col  = pick_col(df, ["Date Of Visit", "Visit Date", "Service Date"], fuzzy_contains=True)
    user_col  = pick_col(df, ["User", "Employee", "Staff"], fuzzy_contains=True)
    client_col= pick_col(df, ["Client", "Patient", "Member"], fuzzy_contains=True)
    hours_col = pick_col(df, ["Hours", "Time"], fuzzy_contains=True)

    missing = [k for k,v in {
        "Date Of Visit": date_col, "User": user_col, "Client": client_col, "Hours": hours_col
    }.items() if v is None]
    if missing:
        raise ValueError(f"Could not find required columns: {', '.join(missing)}")

    # Coerce types
    df = df[[date_col, user_col, client_col, hours_col]].copy()
    df.columns = ["Date Of Visit","User","Client","Hours"]
    df["Date Of Visit"] = pd.to_datetime(df["Date Of Visit"], errors="coerce")
    df["Hours"] = pd.to_numeric(df["Hours"], errors="coerce")
    # Drop rows without dates or hours
    df = df.dropna(subset=["Date Of Visit"])
    df["Hours"] = df["Hours"].fillna(0.0)
    # Clean strings
    df["User"] = df["User"].astype(str).str.strip()
    df["Client"] = df["Client"].astype(str).str.strip()
    return df

def build_sota_tables(df: pd.DataFrame, year: int,
                      include_totals_row: bool = True,
                      include_days_worked_row: bool = False) -> dict:
    """
    From the normalized Smith DF → build a dict of {employee: DataFrame}
    in SOTA layout:
      Client | Week 1 | Week 2 | Week 3 | Week 4 | Week 5 | Month Total
    With Month banner rows separating months.
    """
    work = df.copy()
    work = work[work["Date Of Visit"].dt.year == year]
    if work.empty:
        return {}

    work["Month"] = work["Date Of Visit"].dt.month_name()
    work["MonthNum"] = work["Date Of Visit"].dt.month
    work["Day"]   = work["Date Of Visit"].dt.day
    work["SOTA_Week"] = work["Day"].map(sota_week).astype("Int64")

    # Base aggregation: hours by Employee × Month × Client × SOTA_Week
    agg = (work.groupby(["User","Month","MonthNum","Client","SOTA_Week"], dropna=False)["Hours"]
                .sum().reset_index())
    # For '#days worked' we also want distinct day counts per week
    day_counts = None
    if include_days_worked_row:
        day_counts = (work.groupby(["User","Month","MonthNum","SOTA_Week"])["Date Of Visit"]
                           .nunique().reset_index(name="DaysWorked"))

    # Pivot to wide weeks (1..5)
    def make_employee_sheet(emp):
        emp_df = agg[agg["User"] == emp].copy()
        if emp_df.empty:
            return pd.DataFrame(columns=["Client","Week 1","Week 2","Week 3","Week 4","Week 5","Month Total"])

        # Keep month order Jan..Dec
        month_order = sorted(emp_df["MonthNum"].unique())
        output_rows = []

        for mnum in month_order:
            mon_name = MONTHS[mnum-1]
            mon = emp_df[emp_df["MonthNum"] == mnum].copy()

            # Month banner row
            output_rows.append({
                "Client": mon_name,
                "Week 1": "", "Week 2": "", "Week 3": "", "Week 4": "", "Week 5": "", "Month Total": ""
            })

            # Pivot per client
            pivot = (mon.pivot_table(index="Client", columns="SOTA_Week", values="Hours",
                                     aggfunc="sum", fill_value=0.0)
                        .reindex(columns=[1,2,3,4,5], fill_value=0.0))
            pivot = pivot.rename(columns={1:"Week 1",2:"Week 2",3:"Week 3",4:"Week 4",5:"Week 5"})
            pivot["Month Total"] = pivot[["Week 1","Week 2","Week 3","Week 4","Week 5"]].sum(axis=1)
            pivot = pivot.reset_index()

            # Sort clients: put real client names first (exclude empty/NA)
            pivot["Client_sort"] = pivot["Client"].replace({"nan":"", np.nan:""}).astype(str)
            pivot = pivot.sort_values("Client_sort").drop(columns="Client_sort")

            for _, r in pivot.iterrows():
                output_rows.append({
                    "Client": r["Client"],
                    "Week 1": r.get("Week 1", 0.0),
                    "Week 2": r.get("Week 2", 0.0),
                    "Week 3": r.get("Week 3", 0.0),
                    "Week 4": r.get("Week 4", 0.0),
                    "Week 5": r.get("Week 5", 0.0),
                    "Month Total": r["Month Total"]
                })

            # Optional '#days worked' row (counts, not hours)
            if include_days_worked_row:
                dsub = day_counts[(day_counts["User"] == emp) & (day_counts["MonthNum"] == mnum)]
                w_counts = {f"Week {i}": 0 for i in range(1,6)}
                for _, rw in dsub.iterrows():
                    w_counts[f"Week {int(rw['SOTA_Week'])}"] = int(rw["DaysWorked"])
                output_rows.append({
                    "Client": "#days worked",
                    **w_counts,
                    "Month Total": sum(w_counts.values())
                })

            # Optional 'Totals' row (sum of hours)
            if include_totals_row:
                totals = {"Week 1": 0.0, "Week 2": 0.0, "Week 3": 0.0, "Week 4": 0.0, "Week 5": 0.0}
                if not pivot.empty:
                    for w in totals.keys():
                        totals[w] = float(pivot[w].sum())
                output_rows.append({
                    "Client": "Totals",
                    **totals,
                    "Month Total": sum(totals.values())
                })

        return pd.DataFrame(output_rows, columns=["Client","Week 1","Week 2","Week 3","Week 4","Week 5","Month Total"])

    employees = list(OrderedDict.fromkeys(work["User"].tolist()))
    sheets = {}
    for emp in sorted(set(employees)):
        sheets[emp] = make_employee_sheet(emp)

    return sheets

def build_excel_bytes(sheets_dict: dict) -> bytes:
    """Write {sheet_name: DataFrame} to a single XLSX in memory with some light formatting."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet_name, df in sheets_dict.items():
            # Excel sheet names max length 31
            safe_name = sheet_name[:31] if sheet_name else "Sheet"
            df.to_excel(writer, sheet_name=safe_name, index=False)
            # light formatting
            wb = writer.book
            ws = writer.sheets[safe_name]
            header_fmt = wb.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1})
            num_fmt = wb.add_format({"num_format": "0.00"})
            # Set column widths
            ws.set_column("A:A", 28)
            ws.set_column("B:G", 12)
            # Header style
            for col, _ in enumerate(df.columns):
                ws.write(0, col, df.columns[col], header_fmt)
            # Number format for week/total columns
            for r in range(1, len(df)+1):
                for c in range(1, 7):  # Week1..Week5 + Month Total
                    # keep blank cells as-is
                    val = df.iloc[r-1, c]
                    if isinstance(val, (int, float, np.floating)):
                        ws.write_number(r, c, float(val), num_fmt)
    output.seek(0)
    return output.read()

# ----------------------------
# UI
# ----------------------------
st.title("🗂️ Smith → SOTA Converter")
st.caption("Upload a Smith Integrated Care Services *Report Time* Excel file and download a SOTA-style workbook (one sheet per employee, with Week 1–5).")

with st.sidebar:
    st.header("Options")
    year = st.number_input("Year", min_value=2000, max_value=2100, value=2025, step=1)
    include_totals = st.checkbox("Include 'Totals' row per month", value=True)
    include_days = st.checkbox("Include '#days worked' row (counts, not hours)", value=False)
    st.markdown("---")
    st.markdown("**SOTA week map**  \nW1 = 1–7  \nW2 = 8–14  \nW3 = 15–21  \nW4 = 22–28  \nW5 = 29–end")

uploaded = st.file_uploader("Upload Smith Integrated Excel (.xlsx)", type=["xls","xlsx"])

if uploaded is not None:
    try:
        df = load_smith_dataframe(uploaded)
    except Exception as e:
        st.error(f"Could not parse the uploaded file: {e}")
        st.stop()

    # Quick profile
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.metric("Rows", f"{len(df):,}")
    with c2: st.metric("Employees", df["User"].nunique())
    with c3: st.metric("Clients", df["Client"].nunique())
    with c4: st.metric(f"Hours in {year}", f"{df[df['Date Of Visit'].dt.year==year]['Hours'].sum():,.2f}")

    # Build SOTA sheets
    sheets = build_sota_tables(df, year, include_totals_row=include_totals, include_days_worked_row=include_days)

    if not sheets:
        st.warning(f"No rows found for year {year}. Check your file/year selection.")
    else:
        # Preview first employee sheet
        first_emp = sorted(sheets.keys())[0]
        st.subheader(f"Preview: {first_emp}")
        st.dataframe(sheets[first_emp].head(30), use_container_width=True)

        # Build downloadable workbook
        xlsx_bytes = build_excel_bytes(sheets)
        st.download_button(
            label="⬇️ Download SOTA-formatted workbook",
            data=xlsx_bytes,
            file_name=f"SOTA_{year}_AllEmployees.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Optional: reconciliation (sanity check)
        with st.expander("Reconciliation (Employee × Month) — sanity check", expanded=False):
            # Compute SOTA totals from sheets back into a tidy table
            # (This is purely to help confirm that the transformation didn’t drop anything.)
            tidy = []
            for emp, sdf in sheets.items():
                current_month = None
                for _, row in sdf.iterrows():
                    client = str(row["Client"])
                    if client in MONTHS:
                        current_month = client
                        continue
                    if client.strip().lower() in {"totals", "#days worked"}:
                        continue
                    total = pd.to_numeric(row.get("Month Total", np.nan), errors="coerce")
                    if not pd.isna(total):
                        tidy.append([emp, current_month, float(total)])
            sota_check = pd.DataFrame(tidy, columns=["Employee","Month","Hours"]).groupby(["Employee","Month"])["Hours"].sum().reset_index()

            # Smith side
            dd = df.copy()
            dd["Year"] = dd["Date Of Visit"].dt.year
            dd["Month"] = dd["Date Of Visit"].dt.month_name()
            smith_check = dd[dd["Year"]==year].groupby(["User","Month"])["Hours"].sum().reset_index().rename(columns={"User":"Employee"})

            cmp = pd.merge(
                sota_check.assign(Employee=lambda d: d["Employee"].str.strip()),
                smith_check.assign(Employee=lambda d: d["Employee"].str.strip()),
                on=["Employee","Month"],
                how="outer",
                suffixes=("_SOTA","_Smith")
            ).fillna(0.0)
            cmp["Diff"] = cmp["Hours_SOTA"] - cmp["Hours_Smith"]
            st.dataframe(cmp.sort_values(["Employee","Month"]), use_container_width=True)
else:
    st.info("Upload a Smith Integrated Excel to begin.")

