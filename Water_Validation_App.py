# Water_Validation_App.py
# Streamlit app for Water Quality Data Validation (GENERAL → CORE → ECOLI → ADVANCED → RIPARIAN)
# Now ENFORCES (not just flags):
# - Transparency Tube 0–1.2 m → out-of-range cleared
# - Secchi 0.2–5.0 m (TX QA) & Secchi > Depth → cleared
# - DO titrations diff > 0.5 mg/L → DO values cleared
# - Calibration time >24h (per-probe) → related probe data cleared
# - Conductivity post-cal ±20% → cond cleared
# - E. coli incubation/time/field blank/CFU mismatch/>200 → cleared
# - Contextual outliers >3σ (per-site) → cleared
# Plus: exports Final_Combined.xlsx »

import streamlit as st
import pandas as pd
import numpy as np
import tempfile, zipfile, io, os, re
from datetime import datetime
from typing import Optional, Tuple

from openpyxl import load_workbook  # noqa: F401

st.set_page_config(layout="wide", page_title=" Water Quality Data Validation App")
st.title(" Water Quality Data Validation App")

COND_CANDIDATES = ["Conductivity (µS/cm)", "Conductivity (?S/cm)"]

def save_excel(df: pd.DataFrame, path: str):
    df.to_excel(path, index=False, engine="openpyxl")

def tmp_dir():
    if "tmpdir" not in st.session_state:
        st.session_state.tmpdir = tempfile.mkdtemp(prefix="wqval_")
    return st.session_state.tmpdir

def mark_success(msg):
    st.success("   " + msg)

def path_with_suffix(basename: str, suffix: str):
    d = tmp_dir()
    name, ext = os.path.splitext(basename)
    return os.path.join(d, f"{name}_{suffix}.xlsx")

def init_state():
    for k in [
        "input_basename",
        "df_original",
        "df_general_clean","df_general_annot",
        "df_core_clean","df_core_annot",
        "df_ecoli_clean","df_ecoli_annot",
        "df_adv_clean","df_adv_annot",
        "df_rip_clean","df_rip_annot",
        "df_final_combined","p_final_combined",
        "df_clean_all","p_clean_all",
    ]:
        st.session_state.setdefault(k, None)
init_state()

def first_available(*keys, require_nonempty: bool = False):
    for k in keys:
        df = st.session_state.get(k)
        if isinstance(df, pd.DataFrame):
            if (not require_nonempty) or (not df.empty):
                return df
    return None

def get_cond_col(df: pd.DataFrame) -> Optional[str]:
    return next((c for c in COND_CANDIDATES if c in df.columns), None)

def get_tt_col(df: pd.DataFrame) -> Optional[str]:
    if "Transparency Tube (m)" in df.columns:
        return "Transparency Tube (m)"
    for c in df.columns:
        cl = str(c).lower()
        if "transparency" in cl and "tube" in cl:
            return c
    return None

def get_secchi_col(df: pd.DataFrame) -> Optional[str]:
    candidates = [
        "Secchi Disk Transparency - Average",
        "Secchi Transparency - Average",
        "Secchi Depth (m)",
        "Secchi (m)",
        "Secchi"
    ]
    for c in candidates:
        if c in df.columns:
            return c
    for c in df.columns:
        cl = str(c).lower()
        if "secchi" in cl and ("transpar" in cl or "depth" in cl):
            return c
    return None

def parse_hour_from_time_string(t) -> Optional[int]:
    try:
        s = str(t).strip()
        m = re.match(r"^(\d{1,2})[:\.]?(\d{2})?", s)
        if not m:
            return None
        h = int(m.group(1))
        return h if 0 <= h <= 23 else None
    except:
        return None

def try_parse_datetime(date_val, time_val=None) -> Optional[datetime]:
    try:
        if pd.isna(date_val):
            return None
        if isinstance(date_val, (pd.Timestamp, datetime)):
            dt = pd.to_datetime(date_val).to_pydatetime()
        else:
            dt = pd.to_datetime(date_val, errors="coerce")
            if pd.isna(dt):
                return None
            dt = dt.to_pydatetime()
    except:
        return None
    if time_val is not None:
        try:
            s = str(time_val)
            m = re.match(r"^(\d{1,2}):(\d{2})", s)
            if m:
                h, minute = int(m.group(1)), int(m.group(2))
                if 0 <= h <= 23 and 0 <= minute <= 59:
                    dt = dt.replace(hour=h, minute=minute, second=0, microsecond=0)
                else:
                    dt = dt.replace(hour=12, minute=0, second=0, microsecond=0)
            else:
                h = parse_hour_from_time_string(s)
                if h is not None:
                    dt = dt.replace(hour=h, minute=0, second=0, microsecond=0)
                else:
                    dt = dt.replace(hour=12, minute=0, second=0, microsecond=0)
        except:
            dt = dt.replace(hour=12, minute=0, second=0, microsecond=0)
    return dt

def is_truthy_flag(val) -> bool:
    if pd.isna(val):
        return False
    s = str(val).strip().lower()
    return s in {"y","yes","true","1","flag","flagged","invalid","bad","exclude","remove"}

# ---------- IQR helpers ----------
NON_QUALITY_KEYWORDS = ["Data","Unnamed","Sample","Site","Weather","Notes","Changes","All_","Present"]

def detect_quality_numeric_columns(df: pd.DataFrame, non_quality_keywords=NON_QUALITY_KEYWORDS):
    cols = []
    for c in df.columns:
        name = str(c)
        if not any(kw in name for kw in non_quality_keywords) and pd.api.types.is_numeric_dtype(df[c]):
            cols.append(c)
    return cols

def compute_iqr_bounds(series: pd.Series, k: float = 1.5):
    q1 = series.quantile(0.25)
    q3 = series.quantile(0.75)
    iqr = q3 - q1
    lower = q1 - k * iqr
    upper = q3 + k * iqr
    return q1, q3, lower, upper

def iqr_clean(df: pd.DataFrame, cols: list, k: float = 1.5):
    out = df.copy()
    rows = []
    for col in cols:
        s = pd.to_numeric(out[col], errors="coerce").dropna()
        if s.empty:
            rows.append({"column": col, "Q1": np.nan, "Q3": np.nan,"lower": np.nan,"upper": np.nan,
                         "num_outliers": 0, "num_nonnull": 0, "pct_outliers": 0.0})
            continue
        q1, q3, lower, upper = compute_iqr_bounds(s, k=k)
        mask = (out[col] < lower) | (out[col] > upper)
        n_out = int(mask.sum())
        n_nonnull = int(out[col].notna().sum())
        pct = round((n_out / n_nonnull * 100.0), 3) if n_nonnull else 0.0
        out.loc[mask, col] = np.nan
        rows.append({"column": col, "Q1": q1, "Q3": q3, "lower": lower, "upper": upper,
                     "num_outliers": n_out, "num_nonnull": n_nonnull, "pct_outliers": pct})
    report = pd.DataFrame(rows)
    return out, report

def make_key(df: pd.DataFrame) -> pd.Series:
    cols = []
    if "Group or Affiliation" in df.columns: cols.append(df["Group or Affiliation"].astype(str))
    if "Site ID: Site Name" in df.columns:   cols.append(df["Site ID: Site Name"].astype(str))
    if "Sample Date" in df.columns:          cols.append(pd.to_datetime(df["Sample Date"], errors="coerce").dt.strftime("%Y-%m-%d").fillna(""))
    if "Sample Time Final Format" in df.columns: cols.append(df["Sample Time Final Format"].astype(str))
    if not cols: return df.index.astype(str)
    out = cols[0].fillna("")
    for c in cols[1:]: out = out.str.cat(c.fillna(""), sep="|")
    return out

def build_final_combined(base_df: pd.DataFrame,
                         g_annot: Optional[pd.DataFrame],
                         c_annot: Optional[pd.DataFrame],
                         e_annot: Optional[pd.DataFrame],
                         a_annot: Optional[pd.DataFrame],
                         r_annot: Optional[pd.DataFrame]) -> pd.DataFrame:
    final = base_df.copy()
    final["_key_"] = make_key(final)
    def pick(df, cols):
        if df is None: return None
        use = [c for c in cols if c in df.columns]
        if not use: return None
        tmp = df[use].copy(); tmp["_key_"] = make_key(df); return tmp
    blocks = [
        ("GENERAL_Notes",       pick(g_annot, ["ValidationNotes"])),
        ("GENERAL_Changes",     pick(g_annot, ["TransformNotes"])),
        ("CORE_Notes",          pick(c_annot, ["CORE_Notes"])),
        ("CORE_Changes",        pick(c_annot, ["CORE_ChangeNotes"])),
        ("ECOLI_Notes",         pick(e_annot, ["ECOLI_ValidationNotes"])),
        ("ECOLI_Changes",       pick(e_annot, ["ECOLI_ChangeNotes"])),
        ("ADVANCED_Notes",      pick(a_annot, ["ADVANCED_ValidationNotes"])),
        ("ADVANCED_Changes",    pick(a_annot, ["ADVANCED_ChangeNotes"])),
        ("RIPARIAN_Notes",      pick(r_annot, ["RIPARIAN_ValidationNotes"])),
        ("RIPARIAN_Changes",    pick(r_annot, ["RIPARIAN_ChangeNotes"])),
    ]
    for label, blk in blocks:
        if blk is None: final[label] = ""; continue
        val_cols = [c for c in blk.columns if c != "_key_"]
        if not val_cols: final[label] = ""; continue
        value_col = val_cols[0]
        blk_ren = blk[["_key_", value_col]].rename(columns={value_col: label})
        final = final.merge(blk_ren, on="_key_", how="left"); final[label] = final[label].fillna("")
    def cat_cols(row, cols):
        vals = [str(row[c]).strip() for c in cols if c in row.index and str(row[c]).strip() != ""]
        return " | ".join(vals)
    note_cols = ["GENERAL_Notes","CORE_Notes","ECOLI_Notes","ADVANCED_Notes","RIPARIAN_Notes"]
    chg_cols  = ["GENERAL_Changes","CORE_Changes","ECOLI_Changes","ADVANCED_Changes","RIPARIAN_Changes"]
    final["All_Notes"] = final.apply(lambda r: cat_cols(r, note_cols), axis=1)
    final["All_ChangeNotes"] = final.apply(lambda r: cat_cols(r, chg_cols), axis=1)
    ordered = [c for c in final.columns if c not in ["_key_", "All_Notes", "All_ChangeNotes"]]
    final = final[ordered + ["All_Notes", "All_ChangeNotes"]]
    final.drop(columns=["_key_"], inplace=True, errors="ignore")
    return final

# -------------------- GENERAL --------------------
def run_general(df0: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    df = df0.copy()
    df["ValidationNotes"] = ""
    df["ValidationColorKey"] = ""
    df["TransformNotes"] = ""

    cond_col = get_cond_col(df)
    row_delete_indices = set()

    # Duplicates
    key_cols = [c for c in ["Group or Affiliation","Site ID: Site Name","Sample Date","Sample Time Final Format"] if c in df.columns]
    if len(key_cols) >= 2:
        dup_mask = df.duplicated(subset=key_cols, keep="first")
        df.loc[dup_mask, "ValidationNotes"] += "Duplicate row (same site/date/time); "
        row_delete_indices.update(df[dup_mask].index.tolist())

    # Flagged rows
    flag_cols = [c for c in df.columns if "flag" in c.lower()]
    if flag_cols:
        fl_mask = df[flag_cols].applymap(is_truthy_flag).any(axis=1)
        df.loc[fl_mask, "ValidationNotes"] += "Row flagged by data flag column; "
        row_delete_indices.update(df[fl_mask].index.tolist())

    # Watershed site count >=3
    if "Group or Affiliation" in df.columns and "Site ID: Site Name" in df.columns:
        site_counts = df.groupby("Group or Affiliation")["Site ID: Site Name"].nunique()
        invalid_ws = site_counts[site_counts < 3].index
        mask = df["Group or Affiliation"].isin(invalid_ws)
        df.loc[mask, "ValidationNotes"] += "Less than 3 sites in watershed; "
        df.loc[mask, "ValidationColorKey"] += "watershed_or_events;"
        row_delete_indices.update(df[mask].index.tolist())

    # Site event count >=10
    if "Site ID: Site Name" in df.columns and "Sample Date" in df.columns:
        df["Sample Date"] = pd.to_datetime(df["Sample Date"], errors="coerce")
        event_counts = df.groupby("Site ID: Site Name")["Sample Date"].nunique()
        low_event_sites = event_counts[event_counts < 10].index
        mask = df["Site ID: Site Name"].isin(low_event_sites)
        df.loc[mask, "ValidationNotes"] += "Fewer than 10 events; "
        df.loc[mask, "ValidationColorKey"] += "watershed_or_events;"
        row_delete_indices.update(df[mask].index.tolist())

    # Invalid Sample Date
    if "Sample Date" in df.columns:
        mask = df["Sample Date"].isna()
        df.loc[mask, "ValidationNotes"] += "Missing or invalid Sample Date; "
        df.loc[mask, "ValidationColorKey"] += "time;"
        row_delete_indices.update(df[mask].index.tolist())

    # Sample Time parsing + midday note (note-only)
    if "Sample Time Final Format" in df.columns:
        mask_bad = df["Sample Time Final Format"].apply(lambda t: parse_hour_from_time_string(t) is None)
        df.loc[mask_bad, "ValidationNotes"] += "Unparsable Sample Time; "
        df.loc[mask_bad, "ValidationColorKey"] += "time;"
        row_delete_indices.update(df[mask_bad].index.tolist())
        hours = df["Sample Time Final Format"].apply(parse_hour_from_time_string)
        mask_mid = hours.apply(lambda h: (h is not None) and (12 <= h < 16))
        df.loc[mask_mid, "ValidationNotes"] += "Sample time in 12:00–16:00 window (verify consistency); "

    # Missing all core params → drop row
    core_params = [
        "pH (standard units)",
        "Dissolved Oxygen (mg/L) Average",
        "Water Temperature (° C)",
        cond_col if cond_col else "Conductivity (µS/cm)",
        "Salinity (ppt)",
    ]
    for idx, row in df.iterrows():
        vals = []
        for p in core_params:
            if p in df.columns:
                vals.append(row.get(p))
        if vals and all((pd.isna(v) or v == 0) for v in vals):
            df.at[idx, "ValidationNotes"] += "All core parameters missing or invalid; "
            df.at[idx, "ValidationColorKey"] += "range;"
            row_delete_indices.add(idx)

    # Standard ranges & notes
    standard_ranges = {}; note_texts = {}
    if cond_col:
        standard_ranges[cond_col] = (50, 1500); note_texts[cond_col] = "Conductivity out of range [50–1500]; "
    # Transparency Tube
    tt_col = get_tt_col(df)
    if tt_col and tt_col in df.columns:
        if "(m" not in str(tt_col).lower():
            df["ValidationNotes"] += f"{tt_col} unit not labeled as meters (m); assumed meters; "
            df["ValidationColorKey"] += "unit;"
        df[tt_col] = pd.to_numeric(df[tt_col], errors="coerce")
        standard_ranges[tt_col] = (0.0, 1.2)
        note_texts[tt_col] = "Transparency Tube out of equipment range [0–1.2 m]; "
    # Secchi TX QA
    secchi_col = get_secchi_col(df)
    if secchi_col and secchi_col in df.columns:
        if "(m" not in str(secchi_col).lower():
            df["ValidationNotes"] += f"{secchi_col} unit not labeled as meters (m); assumed meters; "
            df["ValidationColorKey"] += "unit;"
        df[secchi_col] = pd.to_numeric(df[secchi_col], errors="coerce")
        standard_ranges[secchi_col] = (0.2, 5.0)
        note_texts[secchi_col] = "Secchi out of Texas-practical QA window [0.2–5.0 m]; "

    extra = {
        "pH (standard units)": ((6.5, 9.0), "pH out of range [6.5–9.0]; "),
        "Dissolved Oxygen (mg/L) Average": ((5.0, 14.0), "DO out of range [5.0–14.0]; "),
        "Salinity (ppt)": ((0, 35), "Salinity out of range [0–35]; "),
        "Water Temperature (° C)": ((0, 35), "Temp out of range [0–35]; "),
        "Air Temperature (° C)": ((-10, 50), "Air Temp out of range [-10–50]; "),
        "Turbidity": ((0, 1000), "Turbidity out of range [0–1000]; "),
        "E. Coli Average": ((1, 235), "E. Coli out of range [1–235]; "),
        "Nitrate-Nitrogen VALUE (ppm or mg/L)": ((0, 10), "Nitrate out of range [0–10]; "),
        "Orthophosphate": ((0, 0.5), "Orthophosphate out of range [0–0.5]; "),
        "DO (%)": ((80, 120), "DO % out of range [80–120]; "),
        "Total Phosphorus (mg/L)": ((0, 0.05), "TP out of range [0–0.05]; "),
    }
    for k,(rng,txt) in extra.items(): standard_ranges[k] = rng; note_texts[k] = txt

    # Apply ranges → clear cells
    for col, (mn, mx) in standard_ranges.items():
        if col in df.columns:
            vals = pd.to_numeric(df[col], errors="coerce")
            m = (vals < mn) | (vals > mx)
            df.loc[m, "ValidationNotes"] += note_texts[col]
            df.loc[m, "ValidationColorKey"] += "range;"
            df.loc[m, col] = np.nan

    # Contextual outliers (>3σ) → CLEAR (not just flag)
    if "Site ID: Site Name" in df.columns:
        for col in standard_ranges:
            if col in df.columns:
                vals = pd.to_numeric(df[col], errors="coerce")
                means = vals.groupby(df["Site ID: Site Name"]).transform("mean")
                stds  = vals.groupby(df["Site ID: Site Name"]).transform("std")
                z = (vals - means) / stds
                mask = z.abs() > 3
                idxs = mask[mask].index
                df.loc[idxs, "ValidationNotes"] += f"{col} contextual outlier (>3σ) cleared; "
                df.loc[idxs, "ValidationColorKey"] += "contextual_outlier;"
                df.loc[idxs, col] = np.nan

    # Expired reagents → note (clearing specific params بدون نقشه سِت تجهیز ممکن نیست)
    if "Chemical Reagents Used" in df.columns:
        mask = df["Chemical Reagents Used"].astype(str).str.contains("expired", case=False, na=False)
        df.loc[mask, "ValidationNotes"] += "Expired reagents used; "
        df.loc[mask, "ValidationColorKey"] += "expired;"

    # Comments if flagged
    if "Comments" in df.columns:
        empty = df["Comments"].isna() | (df["Comments"].astype(str).str.strip() == "")
        flagged = df["ValidationNotes"] != ""
        mask = flagged & empty
        df.loc[mask, "ValidationNotes"] += "No explanation in Comments; "
        df.loc[mask, "ValidationColorKey"] += "comments;"

    # Clean text 'valid/invalid'
    replaced = df.replace(to_replace=r'(?i)\b(valid|invalid)\b', value='', regex=True)
    changed = replaced != df; df.update(replaced)
    df.loc[changed.any(axis=1), "TransformNotes"] += "Removed 'valid/invalid'; "

    # Sort
    if "Site ID: Site Name" in df.columns and "Sample Date" in df.columns:
        df.sort_values(by=["Site ID: Site Name", "Sample Date"], inplace=True)

    df_clean = df.drop(index=list(row_delete_indices))
    return df_clean, df

# -------------------- CORE --------------------
def run_core(df0: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    df = df0.copy()
    df["CORE_Notes"] = ""
    df["CORE_ChangeNotes"] = ""
    row_delete_indices = set()

    cond_col = get_cond_col(df)

    # Helpers to clear params per calibration issue
    def clear_param_by_name(idx, pname):
        pname = (pname or "").lower()
        if "conductivity" in pname and cond_col and cond_col in df.columns:
            df.at[idx, cond_col] = np.nan
            df.at[idx, "CORE_ChangeNotes"] += "Conductivity cleared due to calibration timing; "
        elif pname == "ph" or "ph" in pname:
            col = "pH (standard units)"
            if col in df.columns:
                df.at[idx, col] = np.nan
                df.at[idx, "CORE_ChangeNotes"] += "pH cleared due to calibration timing; "
        elif "dissolved" in pname or "oxygen" in pname or "do" == pname:
            avg = "Dissolved Oxygen (mg/L) Average"
            do1 = "Dissolved Oxygen (mg/L) 1st titration"
            do2 = "Dissolved Oxygen (mg/L) 2nd titration"
            for c in [avg, do1, do2]:
                if c in df.columns:
                    df.at[idx, c] = np.nan
            df.at[idx, "CORE_ChangeNotes"] += "DO cleared due to calibration timing; "

    # Sample depth 0.3m or mid
    if "Sample Depth (meters)" in df.columns and "Total Depth (meters)" in df.columns:
        for idx, row in df.iterrows():
            sample = row["Sample Depth (meters)"]; total = row["Total Depth (meters)"]
            try:
                if not (np.isclose(sample, 0.3, atol=0.05) or np.isclose(sample, total / 2, atol=0.05)):
                    df.at[idx, "CORE_Notes"] += "Sample Depth not 0.3m or mid-depth; "
            except: pass

    # Depth=0 requires Flow=6 → drop row
    if "Flow Severity" in df.columns and "Total Depth (meters)" in df.columns:
        mask = (df["Total Depth (meters)"] == 0) & (df["Flow Severity"] != 6)
        df.loc[mask, "CORE_Notes"] += "Zero Depth with non-dry flow; "
        row_delete_indices.update(df[mask].index.tolist())

    # DO titrations presence + difference
    do1, do2 = "Dissolved Oxygen (mg/L) 1st titration", "Dissolved Oxygen (mg/L) 2nd titration"
    has1, has2 = (do1 in df.columns), (do2 in df.columns)
    if has1 ^ has2:
        missing = do1 if not has1 else do2
        df["CORE_Notes"] += f"Missing DO titration column: {missing}; "
    if has1 and has2:
        diff = (pd.to_numeric(df[do1], errors="coerce") - pd.to_numeric(df[do2], errors="coerce")).abs()
        mask = diff > 0.5
        df.loc[mask, "CORE_Notes"] += "DO Difference > 0.5; "
        # ENFORCE: clear DO values if mismatch
        for c in [do1, do2, "Dissolved Oxygen (mg/L) Average"]:
            if c in df.columns:
                df.loc[mask, c] = np.nan
        df["CORE_ChangeNotes"] += "Cleared DO (diff>0.5); "

        # Round stored DOs when valid
        df["DO1 Rounded"] = pd.to_numeric(df[do1], errors="coerce").round(1)
        df["DO2 Rounded"] = pd.to_numeric(df[do2], errors="coerce").round(1)
        df["CORE_ChangeNotes"] += "Rounded DO to 0.1; "

    # Secchi: 2 sig figs + remove if > depth
    secchi = get_secchi_col(df) or "Secchi Disk Transparency - Average"
    if secchi in df.columns and "Total Depth (meters)" in df.columns:
        def to_two_sigfigs(x):
            try:
                v = float(x)
                if v == 0 or np.isnan(v): return v
                k = int(np.floor(np.log10(abs(v))))
                return round(v, -k + 1)
            except: return x
        before = df[secchi].copy()
        df[secchi] = pd.to_numeric(df[secchi], errors="coerce").apply(to_two_sigfigs)
        changed = before != df[secchi]
        if changed.any():
            df.loc[changed, "CORE_ChangeNotes"] += "Secchi rounded to 2 significant figures; "
        mask_gt = pd.to_numeric(df[secchi], errors="coerce") > pd.to_numeric(df["Total Depth (meters)"], errors="coerce")
        if mask_gt.any():
            df.loc[mask_gt, "CORE_Notes"] += "Secchi > Depth; "
            df.loc[mask_gt, secchi] = np.nan

    # Conductivity auto-format per guide (format only)
    if cond_col:
        def cond_auto_format(val):
            try:
                v = float(val)
                if v > 100: return float(int(round(v / 10.0)) * 10)
                else:       return float(int(round(v)))
            except: return val
        before = df[cond_col].copy()
        df[cond_col] = df[cond_col].apply(cond_auto_format)
        changed = before != df[cond_col]
        if changed.any():
            df.loc[changed, "CORE_ChangeNotes"] += "Conductivity formatted per guide; "

    # Post-Test Calibration ±20% → ENFORCE clear conductivity
    if "Post-Test Calibration Conductivity" in df.columns and "Standard Value" in df.columns and cond_col:
        post_cal = pd.to_numeric(df["Post-Test Calibration Conductivity"], errors="coerce")
        std_val = pd.to_numeric(df["Standard Value"], errors="coerce")
        valid_cal = (post_cal >= 0.8 * std_val) & (post_cal <= 1.2 * std_val)
        bad = ~valid_cal
        df.loc[bad, "CORE_Notes"] += "Post-Test Calibration outside ±20% of standard; "
        df.loc[bad, cond_col] = np.nan
        df.loc[bad, "CORE_ChangeNotes"] += "Conductivity cleared (calibration ±20% fail); "

    # Calibration time within 24h → ENFORCE clear related params
    pre_time_cols  = [c for c in df.columns if ("pre" in c.lower() and "calibration" in c.lower() and "time" in c.lower())]
    post_time_cols = [c for c in df.columns if ("post" in c.lower() and "calibration" in c.lower() and "time" in c.lower())]
    if "Sample Date" in df.columns:
        for idx, row in df.iterrows():
            samp_dt = try_parse_datetime(row.get("Sample Date"), row.get("Sample Time Final Format"))
            if samp_dt is None: continue
            for c in pre_time_cols + post_time_cols:
                pname = "conductivity" if "conductivity" in c.lower() else ("ph" if "ph" in c.lower() else ("dissolved oxygen" if ("dissolved" in c.lower() or "oxygen" in c.lower() or "do" in c.lower()) else None))
                cal_dt = try_parse_datetime(row.get(c)) or try_parse_datetime(row.get("Sample Date"), row.get(c))
                if cal_dt is not None and abs((samp_dt - cal_dt).total_seconds()) > 24*3600:
                    df.at[idx, "CORE_Notes"] += f"Calibration time >24h from sample ({c}); "
                    clear_param_by_name(idx, pname)

    # pH & Temp rounding (format)
    if "pH (standard units)" in df.columns:
        df["pH Rounded"] = pd.to_numeric(df["pH (standard units)"], errors="coerce").round(1)
        df["CORE_ChangeNotes"] += "Rounded pH to 0.1; "
    if "Water Temperature (° C)" in df.columns:
        df["Water Temp Rounded"] = pd.to_numeric(df["Water Temperature (° C)"], errors="coerce").round(1)
        df["CORE_ChangeNotes"] += "Rounded Water Temp to 0.1; "

    # Salinity display
    if "Salinity (ppt)" in df.columns:
        def fmt_sal(val):
            try:
                if pd.isna(val): return val
                v = float(val)
                return "< 2.0" if v < 2.0 else round(v, 1)
            except: return val
        df["Salinity Formatted"] = df["Salinity (ppt)"].apply(fmt_sal)
        df["CORE_ChangeNotes"] += "Formatted Salinity display; "

    # Numeric format notes
    for col in ["Time Spent Sampling/Traveling", "Roundtrip Distance Traveled"]:
        if col in df.columns:
            mask = ~df[col].apply(lambda x: isinstance(x, (int, float, np.integer, np.floating)))
            df.loc[mask, "CORE_Notes"] += f"{col} format not numeric; "

    df_clean = df.drop(index=row_delete_indices)
    return df_clean, df

# -------------------- ECOLI --------------------
def run_ecoli(df0: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    df = df0.copy()
    df["ECOLI_ValidationNotes"] = ""
    df["ECOLI_ChangeNotes"] = ""

    def clear_ecoli(idx, prefixes=("Sample 1","Sample 2"), clear_avg=True):
        for p in prefixes:
            cfu = f"{p}: Colony Forming Units per 100mL"
            if cfu in df.columns: df.at[idx, cfu] = np.nan
        if clear_avg and "E. Coli Average" in df.columns:
            df.at[idx, "E. Coli Average"] = np.nan

    all_zero_cols = [col for col in df.columns if pd.api.types.is_numeric_dtype(df[col]) and (df[col].fillna(0) == 0).all()]

    # Incubation Temperature 30–36°C → else clear E. coli
    col_temp = "Incubation temperature is 33° C +/- 3° C"
    bad_temp_idx = set()
    if col_temp in df.columns and col_temp not in all_zero_cols:
        df[col_temp] = pd.to_numeric(df[col_temp], errors="coerce")
        mask = (df[col_temp] < 30) | (df[col_temp] > 36)
        df.loc[mask, "ECOLI_ValidationNotes"] += "Incubation temperature not in 30–36°C; "
        bad_temp_idx = set(df[mask].index.tolist())
        for idx in bad_temp_idx: clear_ecoli(idx)

    # Incubation Time 28–31h → else clear
    col_time = "Incubation time is between 28-31 hours"
    bad_time_idx = set()
    if col_time in df.columns and col_time not in all_zero_cols:
        df[col_time] = pd.to_numeric(df[col_time], errors="coerce")
        mask = (df[col_time] < 28) | (df[col_time] > 31)
        df.loc[mask, "ECOLI_ValidationNotes"] += "Incubation time not in 28–31h; "
        bad_time_idx = set(df[mask].index.tolist())
        for idx in bad_time_idx: clear_ecoli(idx)

    # Colonies counted <200
    for i, p in enumerate(["Sample 1","Sample 2"], start=1):
        col_cnt = f"{p}: Colonies Counted"
        if col_cnt in df.columns and col_cnt not in all_zero_cols:
            mask = pd.to_numeric(df[col_cnt], errors="coerce") > 200
            df.loc[mask, "ECOLI_ValidationNotes"] += f"{col_cnt} > 200 colonies; "
            for idx in df[mask].index.tolist():
                clear_ecoli(idx)  # clear both & avg (conservative)

    # Field blank → must be no growth
    col_blank = "No colony growth on Field Blank"
    if col_blank in df.columns and col_blank not in all_zero_cols:
        bad_blank = df[col_blank].astype(str).str.lower().isin(["no", "false", "n"])
        df.loc[bad_blank, "ECOLI_ValidationNotes"] += "Colony growth detected in field blank; "
        for idx in df[bad_blank].index.tolist():
            clear_ecoli(idx)

    # E. Coli average: remove zeros & rounding
    col_ecoli = "E. Coli Average"
    if col_ecoli in df.columns and col_ecoli not in all_zero_cols:
        df[col_ecoli] = pd.to_numeric(df[col_ecoli], errors="coerce")
        mask0 = df[col_ecoli] == 0
        df.loc[mask0, "ECOLI_ValidationNotes"] += "E. coli = 0; "
        df.loc[mask0, col_ecoli] = np.nan

        def round_to_2sf_after_int(n):
            if pd.isna(n): return n
            try:
                n_int = int(round(float(n)))
                if n_int == 0: return 0
                k = int(np.floor(np.log10(abs(n_int))))
                return int(round(n_int, -k + 1))
            except: return n
        df["E. Coli Rounded (int→2SF)"] = df[col_ecoli].apply(round_to_2sf_after_int)
        df["ECOLI_ChangeNotes"] += "Rounded E. coli: nearest int then to 2SF; "

    # CFU formula validation
    def cfu_match(row, prefix):
        try:
            count = row[f"{prefix}: Colonies Counted"]
            dilution = row[f"{prefix}: Dilution Factor (Manual)"]
            volume = row[f"{prefix}: Sample Size (mL)"]
            reported = row[f"{prefix}: Colony Forming Units per 100mL"]
            if any(pd.isna([count, dilution, volume, reported])): return True
            calculated = (count * dilution * 100) / volume
            return abs(calculated - reported) <= 10
        except: return True

    for prefix in ["Sample 1","Sample 2"]:
        cols = [f"{prefix}: Colonies Counted", f"{prefix}: Dilution Factor (Manual)",
                f"{prefix}: Sample Size (mL)", f"{prefix}: Colony Forming Units per 100mL"]
        if all(c in df.columns and c not in all_zero_cols for c in cols):
            valid = df.apply(lambda row: cfu_match(row, prefix), axis=1)
            bad = ~valid
            df.loc[bad, "ECOLI_ValidationNotes"] += f"{prefix} CFU formula mismatch; "
            for idx in df[bad].index.tolist():
                clear_ecoli(idx, prefixes=(prefix,), clear_avg=True)

    df_clean = df[df["ECOLI_ValidationNotes"].str.strip() == ""]
    return df_clean, df

# -------------------- ADVANCED --------------------
def run_adv(df0: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    df = df0.copy()
    df["ADVANCED_ValidationNotes"] = ""
    df["ADVANCED_ChangeNotes"] = ""

    all_zero_cols = [col for col in df.columns if pd.api.types.is_numeric_dtype(df[col]) and (df[col].fillna(0) == 0).all()]
    for c in all_zero_cols:
        df["ADVANCED_ChangeNotes"] += f"Skipped checks for unmeasured parameter: {c}; "

    def log_issue(idx, text): df.at[idx, "ADVANCED_ValidationNotes"] += text + "; "

    # Label sanity notes
    phosphate_cols = [c for c in df.columns if "phosphate" in c.lower() and "value" in c.lower() and c not in all_zero_cols]
    nitrate_cols   = [c for c in df.columns if "nitrate-nitrogen" in c.lower() and "value" in c.lower() and c not in all_zero_cols]
    turbidity_cols = [c for c in df.columns if "turbidity" in c.lower() and "result" in c.lower() and c not in all_zero_cols]
    for c in phosphate_cols:
        if ("mg/l" not in c.lower()) and ("ppm" not in c.lower()):
            for idx in df.index: log_issue(idx, f"{c} not labeled in mg/L or ppm")
    for c in nitrate_cols:
        if ("mg/l" not in c.lower()) and ("ppm" not in c.lower()):
            for idx in df.index: log_issue(idx, f"{c} not labeled in mg/L or ppm")
    for c in turbidity_cols:
        if ("ntu" not in c.lower()) and ("jtu" in c.lower()):
            for idx in df.index: log_issue(idx, f"{c} appears to be in JTU not NTU")

    # Record-level units → ENFORCE clear invalid unit values (if a value column exists)
    unit_col = "ResultMeasure/MeasureUnitCode"; param_col = "CharacteristicName"; value_col = None
    for cand in ["ResultMeasureValue","Result Value","Value","ResultValue"]:
        if cand in df.columns: value_col = cand; break

    if unit_col in df.columns and param_col in df.columns and value_col:
        for idx in df.index:
            p = str(df.at[idx, param_col]).lower(); u = str(df.at[idx, unit_col]).lower()
            bad = False
            if "phosphate" in p and u not in ["mg/l","ppm"]: bad = True
            elif "nitrate" in p and u not in ["mg/l","ppm"]: bad = True
            elif "turbidity" in p and u != "ntu": bad = True
            elif "streamflow" in p and u != "ft2/sec": bad = True
            elif "discharge"  in p and u != "ft2/sec": bad = True
            if bad:
                log_issue(idx, f"{p} unit invalid: {u}; value cleared")
                try: df.at[idx, value_col] = np.nan
                except: pass

    # Discharge formatting (correct values)
    col_discharge = "Discharge Recorded"
    if col_discharge in df.columns and col_discharge not in all_zero_cols:
        def fix_discharge(val):
            try:
                v = float(val)
                if v < 10: new_v = round(v, 1); return new_v, None if abs(v - new_v) < 0.05 else f"{v} → {new_v} (1 dec)"
                else:      new_v = round(v);    return new_v, None if float(v).is_integer() else f"{v} → {new_v} (int)"
            except: return val, "Invalid or non-numeric discharge value"
        for idx in df.index:
            val = df.at[idx, col_discharge]; fixed, issue = fix_discharge(val)
            if issue: log_issue(idx, f"Discharge format issue: {issue}")
            if (fixed is not None) and (fixed != val):
                df.at[idx, col_discharge] = fixed
                df.at[idx, "ADVANCED_ChangeNotes"] += f"Discharge corrected {val} → {fixed}; "

    df_clean = df[df["ADVANCED_ValidationNotes"].str.strip() == ""]
    return df_clean, df

# -------------------- RIPARIAN --------------------
def run_rip(df0: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    df = df0.copy()
    df["RIPARIAN_ValidationNotes"] = ""
    df["RIPARIAN_ChangeNotes"] = ""

    def log_change(idx, msg): df.at[idx, "RIPARIAN_ChangeNotes"] += msg + "; "
    def log_issue(idx, msg):  df.at[idx, "RIPARIAN_ValidationNotes"] += msg + "; "

    indicator_cols = [
        "Energy Dissipation","New Plant Colonization","Stabilizing Vegetation",
        "Age Diversity","Species Diversity","Plant Vigor","Water Storage",
        "Bank/Channel Erosion","Sediment Deposition"
    ]
    available_cols = [c for c in indicator_cols if c in df.columns]

    zeroed_columns = []
    for c in available_cols:
        numeric_col = pd.to_numeric(df[c], errors="coerce").fillna(0)
        if numeric_col.eq(0).all(): zeroed_columns.append(c)
    for c in zeroed_columns:
        df["RIPARIAN_ChangeNotes"] += f"Skipped checks for unmeasured parameter: {c}; "

    if "Bank Evaluated" in df.columns:
        for idx, val in df["Bank Evaluated"].items():
            if pd.isna(val) or str(val).strip() == "":
                log_issue(idx, "Bank evaluation missing")

    for idx, row in df.iterrows():
        missing_count = 0
        for c in available_cols:
            if c in zeroed_columns: continue
            val = row.get(c)
            if pd.isna(val) or str(val).strip() == "":
                comments = str(row.get("Comments", "")).strip().lower()
                if comments in ["", "n/a", "na", "none"]:
                    log_issue(idx, f"{c} missing without explanation")
                else:
                    df.at[idx, c] = np.nan
                missing_count += 1
        if missing_count > 0:
            comments = str(row.get("Comments", "")).strip().lower()
            if comments in ["", "n/a", "na", "none"]:
                log_issue(idx, f"Riparian indicators incomplete: {missing_count} missing")

    image_col = "Image of site was submitted"
    if image_col in df.columns:
        for idx, val in df[image_col].items():
            raw = str(val).strip().lower()
            if raw in ["no","false","n","","nan"]:
                log_issue(idx, "Site image not submitted")
            elif raw in ["yes","true","y"]:
                if str(val).strip() != "Yes":
                    log_change(idx, f"Image value standardized: '{val}' → 'Yes'")
                    df.at[idx, image_col] = "Yes"

    df_clean = df[df["RIPARIAN_ValidationNotes"].str.strip() == ""]
    return df_clean, df

# ==== Tabs ====
tabs = st.tabs([
    " Upload File",
    " GENERAL Validation",
    " CORE Validation",
    " ECOLI Validation",
    " ADVANCED Validation",
    " RIPARIAN Validation",
    " Run All & Exports",
    " Outlier Cleaner (IQR)",
    "Cleaning Guide",
])

# ------------------------ Upload ------------------------
with tabs[0]:
    st.header(" Upload Your Excel File (once)")
    uploaded = st.file_uploader("Upload a .xlsx file", type=["xlsx"])
    if uploaded:
        st.session_state.input_basename = os.path.basename(uploaded.name)
        bytes_data = uploaded.read()
        df0 = pd.read_excel(io.BytesIO(bytes_data), engine="openpyxl")
        st.session_state.df_original = df0.copy()
        mark_success("File loaded. You can proceed to the next tabs or use 'Run All'.")
        st.write("Rows:", len(df0), " | Columns:", len(df0.columns))
        with st.expander("Preview first 20 rows"):
            st.dataframe(df0.head(20))

# ------------------------ GENERAL ------------------------
with tabs[1]:
    st.header(" GENERAL Validation")
    if not isinstance(st.session_state.df_original, pd.DataFrame):
        st.info("Upload a file in the first tab to enable this step.")
    else:
        if st.button("Run GENERAL Validation"):
            g_clean, g_annot = run_general(st.session_state.df_original)
            base = st.session_state.input_basename or "input.xlsx"
            p_clean = path_with_suffix(base, "cleaned_GENERAL")
            p_annot = path_with_suffix(base, "annotated_GENERAL")
            save_excel(g_clean, p_clean); save_excel(g_annot, p_annot)
            st.session_state.df_general_clean = g_clean; st.session_state.df_general_annot = g_annot
            mark_success("GENERAL validation complete.")
            c1, c2 = st.columns(2)
            with c1:
                st.download_button("📥 Download cleaned_GENERAL.xlsx", data=open(p_clean, "rb").read(), file_name="cleaned_GENERAL.xlsx")
            with c2:
                st.download_button("📥 Download annotated_GENERAL.xlsx", data=open(p_annot, "rb").read(), file_name="annotated_GENERAL.xlsx")

# ------------------------ CORE ------------------------
with tabs[2]:
    st.header(" CORE Validation")
    src_core = first_available("df_general_clean")
    if src_core is None:
        st.info("Run GENERAL first (or use Run All).")
    else:
        if st.button("Run CORE Validation"):
            c_clean, c_annot = run_core(src_core)
            base = st.session_state.input_basename or "input.xlsx"
            p_clean = path_with_suffix(base, "cleaned_CORE")
            p_annot = path_with_suffix(base, "annotated_CORE")
            save_excel(c_clean, p_clean); save_excel(c_annot, p_annot)
            st.session_state.df_core_clean = c_clean; st.session_state.df_core_annot = c_annot
            mark_success("CORE validation files generated.")
            c1, c2 = st.columns(2)
            with c1:
                st.download_button("📥 Download cleaned_CORE.xlsx", data=open(p_clean, "rb").read(), file_name="cleaned_CORE.xlsx")
            with c2:
                st.download_button("📥 Download annotated_CORE.xlsx", data=open(p_annot, "rb").read(), file_name="annotated_CORE.xlsx")

# ------------------------ ECOLI ------------------------
with tabs[3]:
    st.header(" ECOLI Validation")
    src_ecoli = first_available("df_general_clean")
    if src_ecoli is None:
        st.info("Run GENERAL first (or use Run All).")
    else:
        if st.button("Run ECOLI Validation"):
            e_clean, e_annot = run_ecoli(src_ecoli)
            base = st.session_state.input_basename or "input.xlsx"
            p_clean = path_with_suffix(base, "cleaned_ECOLI")
            p_annot = path_with_suffix(base, "annotated_ECOLI")
            save_excel(e_clean, p_clean); save_excel(e_annot, p_annot)
            st.session_state.df_ecoli_clean = e_clean; st.session_state.df_ecoli_annot = e_annot
            mark_success("ECOLI validation files generated.")
            c1, c2 = st.columns(2)
            with c1:
                st.download_button("  Download cleaned_ECOLI.xlsx", data=open(p_clean, "rb").read(), file_name="cleaned_ECOLI.xlsx")
            with c2:
                st.download_button("  Download annotated_ECOLI.xlsx", data=open(p_annot, "rb").read(), file_name="annotated_ECOLI.xlsx")

# ------------------------ ADVANCED ------------------------
with tabs[4]:
    st.header(" ADVANCED Validation")
    src_adv = first_available("df_ecoli_clean", "df_general_clean", require_nonempty=False)
    if src_adv is None:
        st.info("Run GENERAL (and optionally ECOLI) first, or use Run All.")
    else:
        if st.button("Run ADVANCED Validation"):
            a_clean, a_annot = run_adv(src_adv)
            base = st.session_state.input_basename or "input.xlsx"
            p_clean = path_with_suffix(base, "cleaned_ADVANCED")
            p_annot = path_with_suffix(base, "annotated_ADVANCED")
            save_excel(a_clean, p_clean); save_excel(a_annot, p_annot)
            st.session_state.df_adv_clean = a_clean; st.session_state.df_adv_annot = a_annot
            mark_success("ADVANCED validation files generated.")
            c1, c2 = st.columns(2)
            with c1:
                st.download_button(" Download cleaned_ADVANCED.xlsx", data=open(p_clean, "rb").read(), file_name="cleaned_ADVANCED.xlsx")
            with c2:
                st.download_button(" Download annotated_ADVANCED.xlsx", data=open(p_annot, "rb").read(), file_name="annotated_ADVANCED.xlsx")

# ------------------------ RIPARIAN ------------------------
with tabs[5]:
    st.header(" RIPARIAN Validation")
    src_rip = first_available("df_adv_clean", "df_general_clean", require_nonempty=False)
    if src_rip is None:
        st.info("Run prior steps (or use Run All).")
    else:
        if st.button("Run RIPARIAN Validation"):
            r_clean, r_annot = run_rip(src_rip)
            base = st.session_state.input_basename or "input.xlsx"
            p_clean = path_with_suffix(base, "cleaned_RIPARIAN")
            p_annot = path_with_suffix(base, "annotated_RIPARIAN")
            save_excel(r_clean, p_clean); save_excel(r_annot, p_annot)
            st.session_state.df_rip_clean = r_clean; st.session_state.df_rip_annot = r_annot
            mark_success("RIPARIAN validation files generated.")
            c1, c2 = st.columns(2)
            with c1:
                st.download_button(" Download cleaned_RIPARIAN.xlsx", data=open(p_clean, "rb").read(), file_name="cleaned_RIPARIAN.xlsx")
            with c2:
                st.download_button(" Download annotated_RIPARIAN.xlsx", data=open(p_annot, "rb").read(), file_name="annotated_RIPARIAN.xlsx")

# ------------------------ RUN ALL & EXPORTS ------------------------
with tabs[6]:
    st.header(" Run All (GENERAL → CORE → ECOLI → ADVANCED → RIPARIAN)")
    st.caption("Final_Combined ")

    if not isinstance(st.session_state.df_original, pd.DataFrame):
        st.info("Upload a file in the first tab.")
    else:
        if st.button("Run All Steps"):
            base = st.session_state.input_basename or "input.xlsx"

            # GENERAL
            g_clean, g_annot = run_general(st.session_state.df_original)
            st.session_state.df_general_clean, st.session_state.df_general_annot = g_clean, g_annot
            p_g_clean = path_with_suffix(base, "cleaned_GENERAL"); p_g_annot = path_with_suffix(base, "annotated_GENERAL")
            save_excel(g_clean, p_g_clean); save_excel(g_annot, p_g_annot)

            # CORE
            c_clean, c_annot = run_core(g_clean)
            st.session_state.df_core_clean, st.session_state.df_core_annot = c_clean, c_annot
            p_c_clean = path_with_suffix(base, "cleaned_CORE"); p_c_annot = path_with_suffix(base, "annotated_CORE")
            save_excel(c_clean, p_c_clean); save_excel(c_annot, p_c_annot)

            # ECOLI
            e_clean, e_annot = run_ecoli(g_clean)
            st.session_state.df_ecoli_clean, st.session_state.df_ecoli_annot = e_clean, e_annot
            p_e_clean = path_with_suffix(base, "cleaned_ECOLI"); p_e_annot = path_with_suffix(base, "annotated_ECOLI")
            save_excel(e_clean, p_e_clean); save_excel(e_annot, p_e_annot)

            # ADVANCED (prefer ECOLI-clean if any)
            a_source = e_clean if not e_clean.empty else g_clean
            a_clean, a_annot = run_adv(a_source)
            st.session_state.df_adv_clean, st.session_state.df_adv_annot = a_clean, a_annot
            p_a_clean = path_with_suffix(base, "cleaned_ADVANCED"); p_a_annot = path_with_suffix(base, "annotated_ADVANCED")
            save_excel(a_clean, p_a_clean); save_excel(a_annot, p_a_annot)

            # RIPARIAN (prefer ADVANCED-clean if any)
            r_source = a_clean if not a_clean.empty else g_clean
            r_clean, r_annot = run_rip(r_source)
            st.session_state.df_rip_clean, st.session_state.df_rip_annot = r_clean, r_annot
            p_r_clean = path_with_suffix(base, "cleaned_RIPARIAN"); p_r_annot = path_with_suffix(base, "annotated_RIPARIAN")
            save_excel(r_clean, p_r_clean); save_excel(r_annot, p_r_annot)

            # Final_Combined (notes merged)
            final_base = r_clean if not r_clean.empty else (a_clean if not a_clean.empty else g_clean)
            df_final = build_final_combined(final_base, g_annot, c_annot, e_annot, a_annot, r_annot)
            p_final = path_with_suffix(base, "Final_Combined"); save_excel(df_final, p_final)
            st.session_state.df_final_combined = df_final.copy(); st.session_state.p_final_combined = p_final

            # Cleaned_AllSteps («داده‌های تمیز.xlsx»)
            df_clean_all = (r_clean if not r_clean.empty else (a_clean if not a_clean.empty else (e_clean if not e_clean.empty else (c_clean if not c_clean.empty else g_clean))))
            p_clean_all = path_with_suffix(base, "Cleaned_AllSteps"); save_excel(df_clean_all, p_clean_all)
            st.session_state.df_clean_all = df_clean_all.copy(); st.session_state.p_clean_all = p_clean_all

            st.success(" All steps completed. Final_Combined + داده‌های تمیز آماده است.")

            c1, c2, c3 = st.columns(3)
            with c1:
                st.download_button(" Download Final_Combined.xlsx", data=open(p_final, "rb").read(), file_name="Final_Combined.xlsx")
            with c2:
                st.download_button(" دانلود «داده‌های تمیز.xlsx»", data=open(p_clean_all, "rb").read(), file_name="داده‌های تمیز.xlsx")
            with c3:
                mem_zip = io.BytesIO()
                with zipfile.ZipFile(mem_zip, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                    for path in [p_g_clean, p_g_annot, p_c_clean, p_c_annot, p_e_clean, p_e_annot, p_a_clean, p_a_annot, p_r_clean, p_r_annot, p_final, p_clean_all]:
                        if os.path.exists(path): zf.write(path, arcname=os.path.basename(path))
                mem_zip.seek(0)
                st.download_button(" Download ALL outputs (ZIP incl. Final_Combined & Cleaned_AllSteps)", data=mem_zip.getvalue(), file_name=f"Validation_Outputs_{datetime.now().strftime('%Y%m%d_%H%M')}.zip", mime="application/zip")

# ------------------------ IQR ------------------------
with tabs[7]:
    st.header("🧹 Outlier Cleaner (IQR)")
    st.caption("Set per-column outliers to NaN using IQR bounds (does not drop rows; metadata stays intact).")
    src = first_available("df_final_combined", "df_adv_clean", "df_ecoli_clean", "df_general_clean", "df_original")
    if src is None:
        st.info("Upload a file in the first tab or run earlier steps.")
    else:
        st.write(f"**Source dataframe:** {len(src)} rows × {len(src.columns)} columns")
        auto_cols = detect_quality_numeric_columns(src)
        with st.expander("Select columns & options", expanded=True):
            incl = st.multiselect("Water-quality numeric columns to clean (IQR):", options=auto_cols, default=auto_cols)
            k = st.slider("IQR Multiplier (k)", min_value=1.0, max_value=3.0, value=1.5, step=0.1)
        if st.button("Run IQR Clean"):
            if not incl: st.warning("Please select at least one numeric column.")
            else:
                cleaned, report = iqr_clean(src, incl, k=k)
                base = st.session_state.input_basename or "input.xlsx"
                p_clean = path_with_suffix(base, f"IQR_NoOutliers_k{str(k).replace('.','_')}"); p_report = os.path.join(tmp_dir(), f"IQR_Report_k{str(k).replace('.','_')}.csv")
                save_excel(cleaned, p_clean); report.to_csv(p_report, index=False)
                st.success(" IQR cleaning done. Downloads below:")
                c1, c2 = st.columns(2)
                with c1:
                    st.download_button(" Download IQR-cleaned Excel", data=open(p_clean, "rb").read(), file_name=os.path.basename(p_clean))
                with c2:
                    st.download_button(" Download IQR bounds/report (CSV)", data=open(p_report, "rb").read(), file_name=os.path.basename(p_report), mime="text/csv")
                with st.expander("Preview (first 20 rows)"): st.dataframe(cleaned.head(20))
                with st.expander("Report preview"): st.dataframe(report)

# ------------------------ Guide ------------------------
with tabs[8]:
    st.header(" Download Data Cleaning Guide")
    st.markdown("Download the official data cleaning and validation guide.")
    guide_filename_on_disk = "Validation_Rules_for_Parameters.pdf"
    if os.path.exists(guide_filename_on_disk):
        with open(guide_filename_on_disk, "rb") as f:
            st.download_button(label=" Download Validation Guide (PDF)", data=f.read(), file_name="Validation_Rules_for_Parameters.pdf", mime="application/pdf")
    else:
        st.info("Place 'Validation_Rules_for_Parameters.pdf' next to the app to enable this download.")
