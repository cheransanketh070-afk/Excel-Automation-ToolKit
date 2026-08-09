import os
import re
import io
import glob
from datetime import datetime
import requests
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# ==========================================
# PAGE CONFIGURATION & STYLING
# ==========================================
st.set_page_config(
    page_title="Excel Automation Toolkit Pro",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    .main-header { font-size: 2.2rem; font-weight: 700; color: #1E293B; margin-bottom: 0.2rem; }
    .sub-header { font-size: 1rem; color: #64748B; margin-bottom: 1.5rem; }
    .stButton>button { background-color: #2563EB; color: white; border-radius: 6px; font-weight: 600; }
    .metric-card { background-color: #F8FAFC; border: 1px solid #E2E8F0; padding: 1rem; border-radius: 8px; text-align: center; }
    .sample-box { background-color: #EFF6FF; border: 1px solid #BFDBFE; padding: 1rem; border-radius: 8px; margin-bottom: 1.5rem; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# LEMON SQUEEZY CONFIGURATION & LICENSE CHECK
# ==========================================
LEMON_SQUEEZY_API_URL = "https://api.lemonsqueezy.com/v1/licenses/activate"
LEMON_SQUEEZY_CHECKOUT_URL = "https://yourstore.lemonsqueezy.com/checkout/buy/YOUR_PRODUCT_ID"
PRIVATE_TEST_KEY = "TEST-GUM-7172504D6E9D"
ALLOWED_SAMPLE_NAMES = ["01_Messy_Customer_Data.xlsx", "02_Dataset_A_Customers.csv"]

def verify_lemon_squeezy_license(license_key):
    """Verifies user subscription key via Lemon Squeezy API or Developer Test Key."""
    if not license_key:
        return False, "No key provided"

    # Private developer test key bypass
    if license_key.strip() == PRIVATE_TEST_KEY:
        return True, "Developer Test Key Active (Unlimited Access)"

    try:
        response = requests.post(
            LEMON_SQUEEZY_API_URL,
            headers={"Accept": "application/json"},
            data={"license_key": license_key.strip()}
        )
        data = response.json()
        if response.status_code == 200 and data.get("activated", False):
            meta = data.get("meta", {})
            return True, f"Subscription Active ({meta.get('customer_email', 'Verified User')})"
        else:
            error_msg = data.get("error", "Invalid or expired license key.")
            return False, error_msg
    except Exception as e:
        return False, f"License server connection error: {str(e)}"

# Initialize session license state
if "is_subscribed" not in st.session_state:
    st.session_state.is_subscribed = False
if "license_msg" not in st.session_state:
    st.session_state.license_msg = "Unlicensed - Sample Mode Active"

def is_allowed_file(file_or_files):
    """Checks if uploaded file is an allowed free sample file or if user has active key."""
    if st.session_state.is_subscribed:
        return True
    
    if not file_or_files:
        return False

    if isinstance(file_or_files, list):
        return all(f.name in ALLOWED_SAMPLE_NAMES for f in file_or_files)
    
    return file_or_files.name in ALLOWED_SAMPLE_NAMES

# ==========================================
# GENERATE IN-MEMORY SAMPLE FILES IF MISSING
# ==========================================
@st.cache_data
def get_sample_excel_bytes():
    if os.path.exists("01_Messy_Customer_Data.xlsx"):
        with open("01_Messy_Customer_Data.xlsx", "rb") as f:
            return f.read()
    
    # Dynamic fallback generator
    data = [
        [" CUST-101 ", " john DOE ", "JOHN.DOE@GMAIL.COM", "1234567890", "2024-01-15", 448.79],
        [" CUST-102 ", "SARAH connor", "invalid_email_here", "1234567890", "02/20/2023", 1123.97],
        [" CUST-103 ", "michael scott ", "m.scott@dundermifflin.com", "+1 (987) 654-3210", "2022/12/05", 950.00],
        [np.nan, np.nan, np.nan, np.nan, np.nan, np.nan],
        [" CUST-104 ", "Pam Beesly", "pam.b@gmail..com", "555-0199", "15-08-2021", 120.50],
        [" CUST-101 ", " john DOE ", "JOHN.DOE@GMAIL.COM", "1234567890", "2024-01-15", 448.79]
    ]
    columns = [" Customer ID ", "Full Name", "Email Address", "Phone Number", "Registration Date", "Account Balance"]
    df = pd.DataFrame(data, columns=columns)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

@st.cache_data
def get_sample_csv_bytes():
    if os.path.exists("02_Dataset_A_Customers.csv"):
        with open("02_Dataset_A_Customers.csv", "rb") as f:
            return f.read()
            
    # Dynamic fallback generator
    data = {
        "CustomerID": ["CUST-101", "CUST-102", "CUST-103", "CUST-104"],
        "CustomerName": ["Acme Corp", "Globex", "Initech", "Umbrella Corp"],
        "Region": ["North", "South", "East", "West"]
    }
    df = pd.DataFrame(data)
    return df.to_csv(index=False).encode('utf-8')

# ==========================================
# HELPER FUNCTIONS - DATA CLEANING
# ==========================================
def clean_phone_number(val):
    if pd.isna(val): return val
    digits = re.sub(r'\D', '', str(val))
    if len(digits) == 10:
        return f"+1 ({digits[:3]}) {digits[3:6]}-{digits[6:]}"
    elif len(digits) == 11 and digits.startswith('1'):
        return f"+1 ({digits[1:4]}) {digits[4:7]}-{digits[7:]}"
    return str(val)

def clean_email_address(val):
    if pd.isna(val): return val
    val = str(val).strip().lower()
    match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', val)
    return match.group(0) if match else np.nan

def clean_dataset(df, opts):
    logs = {"initial_rows": len(df), "duplicates_removed": 0, "blanks_removed": 0, "errors": 0}
    df_clean = df.copy()

    if opts.get('remove_blank_rows'):
        prev = len(df_clean)
        df_clean = df_clean.dropna(how='all')
        logs["blanks_removed"] = prev - len(df_clean)

    if opts.get('trim_spaces'):
        for col in df_clean.select_dtypes(include=['object', 'string']).columns:
            df_clean[col] = df_clean[col].astype(str).str.strip()

    cap_type = opts.get('capitalization')
    if cap_type and cap_type != 'None':
        for col in df_clean.select_dtypes(include=['object', 'string']).columns:
            if cap_type == 'UPPERCASE':
                df_clean[col] = df_clean[col].astype(str).str.upper()
            elif cap_type == 'lowercase':
                df_clean[col] = df_clean[col].astype(str).str.lower()
            elif cap_type == 'Title Case':
                df_clean[col] = df_clean[col].astype(str).str.title()

    if opts.get('clean_phones') and opts.get('phone_cols'):
        for col in opts['phone_cols']:
            df_clean[col] = df_clean[col].apply(clean_phone_number)

    if opts.get('clean_emails') and opts.get('email_cols'):
        for col in opts['email_cols']:
            df_clean[col] = df_clean[col].apply(clean_email_address)

    if opts.get('standardize_dates') and opts.get('date_cols'):
        date_fmt = opts.get('date_format', '%Y-%m-%d')
        for col in opts['date_cols']:
            try:
                df_clean[col] = pd.to_datetime(df_clean[col], errors='coerce').dt.strftime(date_fmt)
            except Exception:
                logs["errors"] += 1

    if opts.get('standardize_cols'):
        df_clean.columns = (
            df_clean.columns.astype(str)
            .str.strip()
            .str.lower()
            .str.replace(r'[^\w\s]', '', regex=True)
            .str.replace(r'\s+', '_', regex=True)
        )

    if opts.get('remove_duplicates'):
        prev = len(df_clean)
        subset = opts.get('duplicate_subset') if opts.get('duplicate_subset') else None
        df_clean = df_clean.drop_duplicates(subset=subset)
        logs["duplicates_removed"] = prev - len(df_clean)

    logs["final_rows"] = len(df_clean)
    return df_clean, logs

# ==========================================
# SIDEBAR NAVIGATION & SUBSCRIPTION GATE
# ==========================================
st.sidebar.title("🛠️ Excel Toolkit Pro")

st.sidebar.markdown("### 🔑 Subscription & Licensing")
input_key = st.sidebar.text_input("Lemon Squeezy Key", type="password", help="Enter your Lemon Squeezy license key")

if st.sidebar.button("Activate License"):
    is_valid, msg = verify_lemon_squeezy_license(input_key)
    st.session_state.is_subscribed = is_valid
    st.session_state.license_msg = msg

if st.session_state.is_subscribed:
    st.sidebar.success(f"🟢 {st.session_state.license_msg}")
else:
    st.sidebar.warning(f"🔴 {st.session_state.license_msg}")
    st.sidebar.markdown(
        f'<a href="{LEMON_SQUEEZY_CHECKOUT_URL}" target="_blank" style="text-decoration:none;">'
        f'<button style="width:100%; background-color:#FFC233; color:#000; padding:8px; border-radius:6px; font-weight:bold; border:none; cursor:pointer; margin-top:5px;">'
        f'💳 Subscribe to Access Toolkit</button></a>',
        unsafe_allow_html=True
    )

st.sidebar.markdown("---")

module = st.sidebar.radio(
    "Select Module",
    [
        "Data Cleaning Studio",
        "File & Batch Operations",
        "Dataset Matching & Compare",
        "Power Tools & Conversion",
        "Interactive Dashboard Studio"
    ]
)

st.sidebar.markdown("---")
st.sidebar.caption("Excel Automation Toolkit v1.2 | Free Sample Mode Enabled")

# ==========================================
# FREE SAMPLE FILE DOWNLOAD BANNER
# ==========================================
st.markdown("""
<div class="sample-box">
    <h4>📁 Test the Toolkit for Free</h4>
    <p style="margin-bottom:0.8rem; color:#475569;">No subscription? Download these sample files to test all tools. Custom file processing requires an active subscription.</p>
</div>
""", unsafe_allow_html=True)

sc1, sc2, _ = st.columns([1, 1, 2])

sc1.download_button(
    label="📥 Download Sample Excel Data",
    data=get_sample_excel_bytes(),
    file_name="01_Messy_Customer_Data.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

sc2.download_button(
    label="📥 Download Sample CSV Data",
    data=get_sample_csv_bytes(),
    file_name="02_Dataset_A_Customers.csv",
    mime="text/csv"
)

st.markdown("---")

# ==========================================
# MODULE 1: DATA CLEANING STUDIO
# ==========================================
if module == "Data Cleaning Studio":
    st.markdown('<div class="main-header">Data Cleaning Studio</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Automate repetitive scrubbing, formatting, and standardizing tasks in seconds.</div>', unsafe_allow_html=True)

    uploaded_file = st.file_uploader("Upload Excel or CSV file", type=["xlsx", "xls", "csv"])

    if uploaded_file:
        file_ext = uploaded_file.name.split('.')[-1].lower()
        if file_ext == 'csv':
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        st.subheader("Data Preview")
        st.dataframe(df.head(5), use_container_width=True)

        col1, col2, col3 = st.columns(3)

        with col1:
            st.markdown("### Basic Cleaning")
            rem_dupes = st.checkbox("Remove Duplicate Records", value=True)
            rem_blanks = st.checkbox("Remove Entirely Blank Rows", value=True)
            trim_sp = st.checkbox("Trim Unnecessary Spaces", value=True)
            std_cols = st.checkbox("Standardize Column Names (snake_case)", value=False)

        with col2:
            st.markdown("### Text & Capitalization")
            cap_type = st.selectbox("Capitalization Standard", ["None", "UPPERCASE", "lowercase", "Title Case"])
            clean_em = st.checkbox("Scrub Email Addresses")
            email_cols = st.multiselect("Select Email Columns", df.columns) if clean_em else []

            clean_ph = st.checkbox("Scrub & Format Phone Numbers")
            phone_cols = st.multiselect("Select Phone Columns", df.columns) if clean_ph else []

        with col3:
            st.markdown("### Dates & Missing Values")
            std_dt = st.checkbox("Standardize Date Formats")
            date_cols = st.multiselect("Select Date Columns", df.columns) if std_dt else []
            dt_fmt = st.selectbox("Output Date Format", ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"]) if std_dt else "%Y-%m-%d"

            dupe_sub = st.multiselect("Check Duplicates Based On Specific Columns", df.columns) if rem_dupes else []

        if st.button("⚡ Run Cleaning Pipeline", use_container_width=True):
            if not is_allowed_file(uploaded_file):
                st.error(f"❌ '{uploaded_file.name}' is blocked in free mode. Free users can only process allowed sample files ('01_Messy_Customer_Data.xlsx' or '02_Dataset_A_Customers.csv'). Please subscribe to process custom files.")
            else:
                options = {
                    'remove_duplicates': rem_dupes,
                    'remove_blank_rows': rem_blanks,
                    'trim_spaces': trim_sp,
                    'standardize_cols': std_cols,
                    'capitalization': cap_type,
                    'clean_emails': clean_em,
                    'email_cols': email_cols,
                    'clean_phones': clean_ph,
                    'phone_cols': phone_cols,
                    'standardize_dates': std_dt,
                    'date_cols': date_cols,
                    'date_format': dt_fmt,
                    'duplicate_subset': dupe_sub
                }

                cleaned_df, report = clean_dataset(df, options)
                st.success("Cleaning Pipeline Executed Successfully!")

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Original Rows", report["initial_rows"])
                m2.metric("Cleaned Rows", report["final_rows"])
                m3.metric("Duplicates Removed", report["duplicates_removed"])
                m4.metric("Blank Rows Removed", report["blanks_removed"])

                st.subheader("Cleaned Dataset Preview")
                st.dataframe(cleaned_df.head(10), use_container_width=True)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    cleaned_df.to_excel(writer, index=False, sheet_name='Cleaned_Data')
                processed_data = output.getvalue()

                out_name = f"Cleaned_{uploaded_file.name.split('.')[0]}.xlsx"
                st.download_button(
                    label="📥 Download Cleaned Excel File",
                    data=processed_data,
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# ==========================================
# MODULE 2: FILE & BATCH OPERATIONS
# ==========================================
elif module == "File & Batch Operations":
    st.markdown('<div class="main-header">File & Batch Operations</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Merge, split, extract, and convert multi-file datasets automatically.</div>', unsafe_allow_html=True)

    action = st.selectbox(
        "Select Operation",
        [
            "Merge Multiple Excel/CSV Files",
            "Split Workbook by Column Value",
            "Split Workbook by Year",
            "Batch CSV → Excel Converter",
            "Extract Specific Columns"
        ]
    )

    if action == "Merge Multiple Excel/CSV Files":
        files = st.file_uploader("Upload Files to Combine", type=["csv", "xlsx", "xls"], accept_multiple_files=True)
        if files:
            if st.button("Combine Files"):
                if not is_allowed_file(files):
                    st.error("❌ Custom multi-file merges require an active subscription. Only sample test files can be processed without a key.")
                else:
                    dfs = []
                    for f in files:
                        ext = f.name.split('.')[-1].lower()
                        df_temp = pd.read_csv(f) if ext == 'csv' else pd.read_excel(f)
                        df_temp['Source_File'] = f.name
                        dfs.append(df_temp)
                    merged_df = pd.concat(dfs, ignore_index=True)
                    st.success(f"Successfully merged {len(files)} files into {len(merged_df)} records!")

                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        merged_df.to_excel(writer, index=False, sheet_name='Merged_Data')
                    st.download_button("📥 Download Merged File", data=output.getvalue(), file_name="Merged_Dataset.xlsx")

    elif action == "Split Workbook by Column Value":
        uploaded_file = st.file_uploader("Upload File to Split", type=["xlsx", "csv"])
        if uploaded_file:
            df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
            split_col = st.selectbox("Select Column to Split By", df.columns)

            if st.button("Generate Split Files"):
                if not is_allowed_file(uploaded_file):
                    st.error("❌ Custom file splitting requires an active subscription. Please upload an allowed sample file or subscribe.")
                else:
                    unique_vals = df[split_col].dropna().unique()
                    st.info(f"Generating {len(unique_vals)} split outputs...")

                    import zipfile
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                        for val in unique_vals:
                            sub_df = df[df[split_col] == val]
                            out_b = io.BytesIO()
                            with pd.ExcelWriter(out_b, engine='openpyxl') as writer:
                                sub_df.to_excel(writer, index=False, sheet_name=str(val)[:30])
                            clean_val_str = re.sub(r'[^\w\-_\. ]', '_', str(val))
                            zip_file.writestr(f"Split_{split_col}_{clean_val_str}.xlsx", out_b.getvalue())

                    st.success(f"Split complete! {len(unique_vals)} files packaged.")
                    st.download_button("📥 Download All Split Files (.zip)", data=zip_buffer.getvalue(), file_name="Split_Files.zip")

    elif action == "Split Workbook by Year":
        uploaded_file = st.file_uploader("Upload File", type=["xlsx", "csv"])
        if uploaded_file:
            df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
            date_col = st.selectbox("Select Date Column", df.columns)

            if st.button("Split by Year"):
                if not is_allowed_file(uploaded_file):
                    st.error("❌ Custom file splitting requires an active subscription.")
                else:
                    df['__Temp_Year'] = pd.to_datetime(df[date_col], errors='coerce').dt.year
                    years = df['__Temp_Year'].dropna().unique()

                    import zipfile
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                        for y in years:
                            sub_df = df[df['__Temp_Year'] == y].drop(columns=['__Temp_Year'])
                            out_b = io.BytesIO()
                            with pd.ExcelWriter(out_b, engine='openpyxl') as writer:
                                sub_df.to_excel(writer, index=False, sheet_name=str(int(y)))
                            zip_file.writestr(f"Data_Year_{int(y)}.xlsx", out_b.getvalue())

                    st.download_button("📥 Download Year-Split Files (.zip)", data=zip_buffer.getvalue(), file_name="Split_By_Year.zip")

    elif action == "Batch CSV → Excel Converter":
        files = st.file_uploader("Upload CSV Files", type=["csv"], accept_multiple_files=True)
        if files and st.button("Convert to Excel"):
            if not is_allowed_file(files):
                st.error("❌ Batch CSV conversion for custom files requires an active subscription.")
            else:
                import zipfile
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    for f in files:
                        df = pd.read_csv(f)
                        out_b = io.BytesIO()
                        with pd.ExcelWriter(out_b, engine='openpyxl') as writer:
                            df.to_excel(writer, index=False, sheet_name='Data')
                        base_name = f.name.rsplit('.', 1)[0]
                        zip_file.writestr(f"{base_name}.xlsx", out_b.getvalue())
                st.download_button("📥 Download Converted Excel Files (.zip)", data=zip_buffer.getvalue(), file_name="Converted_Excel_Files.zip")

    elif action == "Extract Specific Columns":
        uploaded_file = st.file_uploader("Upload File", type=["xlsx", "csv"])
        if uploaded_file:
            df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
            selected_cols = st.multiselect("Select Columns to Extract", df.columns)
            if selected_cols and st.button("Extract Data"):
                if not is_allowed_file(uploaded_file):
                    st.error("❌ Extracting columns from custom files requires an active subscription.")
                else:
                    extracted_df = df[selected_cols]
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        extracted_df.to_excel(writer, index=False)
                    st.download_button("📥 Download Extracted Columns", data=output.getvalue(), file_name="Extracted_Columns.xlsx")

# ==========================================
# MODULE 3: DATASET MATCHING & COMPARE
# ==========================================
elif module == "Dataset Matching & Compare":
    st.markdown('<div class="main-header">Dataset Matching & Sheet Comparison</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Reconcile datasets, find missing records, and spot cross-file duplicates.</div>', unsafe_allow_html=True)

    mode = st.radio("Select Workflow", ["Match Two Datasets (VLOOKUP Replacement)", "Compare Two Excel Files / Sheets"])

    if mode == "Match Two Datasets (VLOOKUP Replacement)":
        c1, c2 = st.columns(2)
        with c1:
            f1 = st.file_uploader("Upload Primary File (Dataset A)", type=["xlsx", "csv"], key="m_f1")
        with c2:
            f2 = st.file_uploader("Upload Secondary File (Dataset B)", type=["xlsx", "csv"], key="m_f2")

        if f1 and f2:
            df1 = pd.read_csv(f1) if f1.name.endswith('.csv') else pd.read_excel(f1)
            df2 = pd.read_csv(f2) if f2.name.endswith('.csv') else pd.read_excel(f2)

            col_a, col_b = st.columns(2)
            key1 = col_a.selectbox("Matching Key column in Dataset A", df1.columns)
            key2 = col_b.selectbox("Matching Key column in Dataset B", df2.columns)

            join_type = st.selectbox("Match Type", ["Inner Join (Records in Both)", "Left Join (Keep All Dataset A)", "Find Missing in Dataset B"])

            if st.button("Execute Match"):
                if not (is_allowed_file(f1) and is_allowed_file(f2)):
                    st.error("❌ Matching custom datasets requires an active subscription. Please subscribe to execute custom file matches.")
                else:
                    if join_type == "Inner Join (Records in Both)":
                        res = pd.merge(df1, df2, left_on=key1, right_on=key2, how='inner', suffixes=('_A', '_B'))
                    elif join_type == "Left Join (Keep All Dataset A)":
                        res = pd.merge(df1, df2, left_on=key1, right_on=key2, how='left', suffixes=('_A', '_B'))
                    else:
                        res = df1[~df1[key1].isin(df2[key2])]

                    st.success(f"Matched {len(res)} total records.")
                    st.dataframe(res.head(10), use_container_width=True)

                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        res.to_excel(writer, index=False)
                    st.download_button("📥 Download Match Report", data=output.getvalue(), file_name="Matched_Output.xlsx")

    elif mode == "Compare Two Excel Files / Sheets":
        c1, c2 = st.columns(2)
        with c1:
            f1 = st.file_uploader("Upload File A (Original)", type=["xlsx", "csv"], key="c_f1")
        with c2:
            f2 = st.file_uploader("Upload File B (Modified)", type=["xlsx", "csv"], key="c_f2")

        if f1 and f2:
            df1 = pd.read_csv(f1) if f1.name.endswith('.csv') else pd.read_excel(f1)
            df2 = pd.read_csv(f2) if f2.name.endswith('.csv') else pd.read_excel(f2)

            if st.button("Compare Sheets"):
                if not (is_allowed_file(f1) and is_allowed_file(f2)):
                    st.error("❌ Comparing custom files requires an active subscription.")
                else:
                    common_cols = list(set(df1.columns).intersection(set(df2.columns)))
                    df1_c = df1[common_cols]
                    df2_c = df2[common_cols]

                    diff = df1_c.compare(df2_c)
                    if diff.empty:
                        st.success("The files are identical across common columns!")
                    else:
                        st.warning(f"Found {len(diff)} cell differences between files.")
                        st.dataframe(diff, use_container_width=True)

                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            diff.to_excel(writer)
                        st.download_button("📥 Download Difference Log", data=output.getvalue(), file_name="Comparison_Diff_Log.xlsx")

# ==========================================
# MODULE 4: POWER TOOLS & CONVERSION
# ==========================================
elif module == "Power Tools & Conversion":
    st.markdown('<div class="main-header">Power Tools & Advanced Operations</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Filter, sort, and batch rename worksheets inside workbooks.</div>', unsafe_allow_html=True)

    p_action = st.selectbox("Select Tool", ["Filter & Sort Records", "Batch Rename Worksheets"])

    if p_action == "Filter & Sort Records":
        uploaded_file = st.file_uploader("Upload File", type=["xlsx", "csv"])
        if uploaded_file:
            df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)

            c1, c2 = st.columns(2)
            filter_col = c1.selectbox("Filter Column", ["None"] + list(df.columns))
            sort_col = c2.selectbox("Sort Column", ["None"] + list(df.columns))

            filtered_df = df.copy()

            if filter_col != "None":
                val = st.text_input(f"Filter value for '{filter_col}' (contains text)")
                if val:
                    filtered_df = filtered_df[filtered_df[filter_col].astype(str).str.contains(val, case=False, na=False)]

            if sort_col != "None":
                ascending = st.checkbox("Ascending Order", value=True)
                filtered_df = filtered_df.sort_values(by=sort_col, ascending=ascending)

            st.dataframe(filtered_df.head(10), use_container_width=True)

            if st.button("Download Result"):
                if not is_allowed_file(uploaded_file):
                    st.error("❌ Exporting processed custom data requires an active subscription.")
                else:
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        filtered_df.to_excel(writer, index=False)
                    st.download_button("📥 Download Filtered/Sorted Data", data=output.getvalue(), file_name="Filtered_Sorted_Data.xlsx")

    elif p_action == "Batch Rename Worksheets":
        uploaded_file = st.file_uploader("Upload Multi-Sheet Excel File", type=["xlsx"])
        if uploaded_file:
            xl = pd.ExcelFile(uploaded_file)
            sheet_names = xl.sheet_names
            st.write("Current Sheets:", sheet_names)

            prefix = st.text_input("Add Prefix to All Sheets", value="")
            suffix = st.text_input("Add Suffix to All Sheets", value="")

            if st.button("Rename Sheets"):
                if not is_allowed_file(uploaded_file):
                    st.error("❌ Renaming worksheets in custom files requires an active subscription.")
                else:
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        for s in sheet_names:
                            df_s = pd.read_excel(uploaded_file, sheet_name=s)
                            new_name = f"{prefix}{s}{suffix}"[:31]
                            df_s.to_excel(writer, sheet_name=new_name, index=False)
                    st.success("Worksheets renamed successfully!")
                    st.download_button("📥 Download Renamed Workbook", data=output.getvalue(), file_name="Renamed_Sheets.xlsx")

# ==========================================
# MODULE 5: INTERACTIVE DASHBOARD STUDIO
# ==========================================
elif module == "Interactive Dashboard Studio":
    st.markdown('<div class="main-header">Interactive Dashboard Studio</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">Transform datasets into clean, executive-ready analytical dashboards automatically.</div>', unsafe_allow_html=True)

    uploaded_file = st.file_uploader("Upload Excel or CSV file for Dashboard Generation", type=["xlsx", "csv"], key="dash_file")

    if uploaded_file:
        df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
        
        # Guardrail subscription check
        if not is_allowed_file(uploaded_file):
            st.error(f"❌ Dashboard generation for custom file '{uploaded_file.name}' requires an active subscription. Free users can test this module using sample files.")
        else:
            st.markdown("### 📊 Executive Overview KPI Metrics")
            num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
            cat_cols = df.select_dtypes(include=['object', 'string', 'category']).columns.tolist()

            kpi1, kpi2, kpi3, kpi4 = st.columns(4)
            kpi1.metric("Total Records", f"{len(df):,}")
            kpi2.metric("Total Attributes", f"{len(df.columns)}")

            if num_cols:
                metric_col = kpi3.selectbox("Primary Metric", num_cols, index=0)
                kpi3.metric(f"Total {metric_col}", f"{df[metric_col].sum():,.2f}")
                kpi4.metric(f"Average {metric_col}", f"{df[metric_col].mean():,.2f}")
            else:
                kpi3.metric("Numeric Fields", "None Found")
                kpi4.metric("Categorical Fields", len(cat_cols))

            st.markdown("---")
            st.markdown("### 📈 Visual Analytics Builder")

            dash_type = st.selectbox(
                "Select Visualization Type",
                [
                    "Bar Chart (Category Aggregation)",
                    "Line Chart (Trend / Time Series)",
                    "Scatter Plot (Correlation Analysis)",
                    "Pie / Donut Chart (Distribution)",
                    "Histogram (Data Distribution)",
                    "Heatmap (Correlation Matrix)"
                ]
            )

            c1, c2 = st.columns(2)

            if dash_type == "Bar Chart (Category Aggregation)":
                if cat_cols and num_cols:
                    x_col = c1.selectbox("Category Column (X-Axis)", cat_cols)
                    y_col = c2.selectbox("Numeric Value Column (Y-Axis)", num_cols)
                    agg_func = st.selectbox("Aggregation Function", ["Sum", "Mean", "Count"])
                    
                    if agg_func == "Sum":
                        grouped_df = df.groupby(x_col)[y_col].sum().reset_index()
                    elif agg_func == "Mean":
                        grouped_df = df.groupby(x_col)[y_col].mean().reset_index()
                    else:
                        grouped_df = df.groupby(x_col)[y_col].count().reset_index()

                    fig = px.bar(grouped_df, x=x_col, y=y_col, title=f"{agg_func} of {y_col} by {x_col}", template="plotly_white", text_auto='.2s')
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.warning("Bar charts require at least one categorical column and one numeric column.")

            elif dash_type == "Line Chart (Trend / Time Series)":
                date_or_seq = list(df.columns)
                if num_cols:
                    x_col = c1.selectbox("Time or Sequence Axis (X-Axis)", date_or_seq)
                    y_col = c2.selectbox("Metric Column (Y-Axis)", num_cols)
                    fig = px.line(df, x=x_col, y=y_col, title=f"{y_col} Trend Over {x_col}", template="plotly_white", markers=True)
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.warning("Line charts require at least one numeric metric column.")

            elif dash_type == "Scatter Plot (Correlation Analysis)":
                if len(num_cols) >= 2:
                    x_col = c1.selectbox("X-Axis Variable", num_cols, index=0)
                    y_col = c2.selectbox("Y-Axis Variable", num_cols, index=min(1, len(num_cols)-1))
                    color_col = st.selectbox("Group / Color By (Optional)", ["None"] + cat_cols)
                    
                    color_param = None if color_col == "None" else color_col
                    fig = px.scatter(df, x=x_col, y=y_col, color=color_param, title=f"Correlation: {x_col} vs {y_col}", template="plotly_white")
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.warning("Scatter plots require at least two numeric columns.")

            elif dash_type == "Pie / Donut Chart (Distribution)":
                if cat_cols:
                    names_col = c1.selectbox("Category Field", cat_cols)
                    values_col = c2.selectbox("Value Field (Optional Count if None)", ["Record Count"] + num_cols)
                    
                    hole_val = 0.4 if st.checkbox("Render as Donut Chart", value=True) else 0.0
                    
                    if values_col == "Record Count":
                        fig = px.pie(df, names=names_col, title=f"Distribution by {names_col}", hole=hole_val, template="plotly_white")
                    else:
                        fig = px.pie(df, names=names_col, values=values_col, title=f"{values_col} Share by {names_col}", hole=hole_val, template="plotly_white")
                    
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.warning("Pie charts require at least one categorical column.")

            elif dash_type == "Histogram (Data Distribution)":
                if num_cols:
                    num_col = c1.selectbox("Numeric Field", num_cols)
                    bins = c2.slider("Number of Bins", min_value=5, max_value=50, value=15)
                    fig = px.histogram(df, x=num_col, nbins=bins, title=f"Distribution Frequency of {num_col}", template="plotly_white")
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.warning("Histograms require at least one numeric column.")

            elif dash_type == "Heatmap (Correlation Matrix)":
                if len(num_cols) >= 2:
                    corr = df[num_cols].corr()
                    fig = px.imshow(corr, text_auto=True, aspect="auto", title="Numeric Feature Correlation Heatmap", color_continuous_scale="Blues", template="plotly_white")
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.warning("Correlation Heatmaps require at least two numeric columns.")