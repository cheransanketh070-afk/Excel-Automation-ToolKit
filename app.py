import os
import re
import io
import glob
import zipfile
from datetime import datetime, timedelta
import requests
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# ==========================================
# PAGE CONFIGURATION & MODERN SAAS STYLING
# ==========================================
st.set_page_config(
    page_title="Excel Automation Toolkit Pro",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom High-End SaaS & Sidebar Styling Injection
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }
    
    .main {
        background-color: #F8FAFC;
    }
    
    /* Executive Hero Header */
    .hero-container {
        background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
        padding: 2.2rem 2.5rem;
        border-radius: 16px;
        color: #FFFFFF;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.15);
        margin-bottom: 2rem;
    }
    
    .hero-title {
        font-size: 2.2rem;
        font-weight: 700;
        letter-spacing: -0.02em;
        margin-bottom: 0.4rem;
        color: #FFFFFF;
    }
    
    .hero-subtitle {
        font-size: 1.05rem;
        color: #94A3B8;
        font-weight: 400;
        margin-bottom: 0px;
    }

    /* Professional Sidebar Styling */
    section[data-testid="stSidebar"] {
        background-color: #0F172A !important;
        border-right: 1px solid #1E293B;
    }

    section[data-testid="stSidebar"] * {
        color: #F1F5F9 !important;
    }

    .sidebar-header {
        font-size: 1.3rem;
        font-weight: 700;
        letter-spacing: -0.01em;
        color: #FFFFFF !important;
        margin-bottom: 0.2rem;
    }

    .sidebar-badge {
        display: inline-block;
        background: rgba(37, 99, 235, 0.2);
        border: 1px solid rgba(59, 130, 246, 0.4);
        color: #60A5FA !important;
        font-size: 0.72rem;
        font-weight: 600;
        padding: 2px 8px;
        border-radius: 20px;
        margin-bottom: 1rem;
    }

    .sidebar-status-box {
        background: rgba(30, 41, 59, 0.7);
        border: 1px solid #334155;
        border-radius: 10px;
        padding: 0.85rem 1rem;
        margin-top: 0.75rem;
        margin-bottom: 1.25rem;
    }

    .sidebar-status-title {
        font-size: 0.75rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #94A3B8 !important;
        margin-bottom: 0.25rem;
    }

    .sidebar-status-val {
        font-size: 0.875rem;
        font-weight: 600;
    }

    /* KPI Cards */
    .kpi-card-container {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 12px;
        padding: 1.25rem 1.5rem;
        box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.05);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .kpi-card-container:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.08);
    }
    .kpi-title {
        font-size: 0.825rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #64748B;
        margin-bottom: 0.5rem;
    }
    .kpi-val {
        font-size: 1.8rem;
        font-weight: 700;
        color: #0F172A;
    }

    .free-banner {
        background: linear-gradient(90deg, #EFF6FF 0%, #DBEAFE 100%);
        border: 1px solid #BFDBFE;
        border-radius: 12px;
        padding: 1.25rem;
        margin-bottom: 2rem;
    }
    
    .stButton>button {
        background-color: #2563EB;
        color: #FFFFFF !important;
        font-weight: 600;
        border-radius: 8px;
        padding: 0.5rem 1.25rem;
        border: none;
        box-shadow: 0 2px 4px rgba(37, 99, 235, 0.2);
        transition: all 0.2s ease;
    }
    .stButton>button:hover {
        background-color: #1D4ED8;
        box-shadow: 0 4px 8px rgba(37, 99, 235, 0.3);
    }
    
    .block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# LEMON SQUEEZY CONFIGURATION & LICENSE CHECK
# ==========================================
LEMON_SQUEEZY_API_URL = "https://api.lemonsqueezy.com/v1/licenses/activate"
LEMON_SQUEEZY_CHECKOUT_URL = "https://yourstore.lemonsqueezy.com/checkout/buy/YOUR_PRODUCT_ID"

ALLOWED_SAMPLE_NAMES = [
    "Enterprise_Global_Sales_2026.xlsx",
    "Enterprise_Customer_Master_2026.csv"
]

def verify_lemon_squeezy_license(license_key):
    """Verifies user subscription key via Lemon Squeezy API."""
    if not license_key:
        return False, "No license key entered"

    try:
        response = requests.post(
            LEMON_SQUEEZY_API_URL,
            headers={"Accept": "application/json"},
            data={"license_key": license_key.strip()}
        )
        data = response.json()
        if response.status_code == 200 and data.get("activated", False):
            meta = data.get("meta", {})
            return True, f"License Verified ({meta.get('customer_email', 'Active Pro')})"
        else:
            error_msg = data.get("error", "Invalid or expired license key.")
            return False, error_msg
    except Exception as e:
        return False, f"License verification error: {str(e)}"

# Initialize Session State
if "is_subscribed" not in st.session_state:
    st.session_state.is_subscribed = False
if "license_msg" not in st.session_state:
    st.session_state.license_msg = "Free Tier - Sample Datasets Only"

def is_allowed_file(file_or_files):
    """Enforces subscription requirement for custom uploaded files."""
    if st.session_state.is_subscribed:
        return True
    
    if not file_or_files:
        return False

    if isinstance(file_or_files, list):
        return all(f.name in ALLOWED_SAMPLE_NAMES for f in file_or_files)
    
    return file_or_files.name in ALLOWED_SAMPLE_NAMES

# ==========================================
# HIGH-QUALITY LARGE SAMPLE DATASET GENERATORS
# ==========================================
@st.cache_data
def generate_large_excel_sample():
    """Generates a rich 1,000+ record dataset with 15 numeric/categorical/text columns."""
    np.random.seed(42)
    n_rows = 1050

    names = ["John Doe", "Sarah Connor", "Michael Scott", "Pam Beesly", "Dwight Schrute", 
             "Jim Halpert", "Stanley Hudson", "Phyllis Vance", "Angela Martin", "Kevin Malone",
             "Rachel Green", "Ross Geller", "Chandler Bing", "Monica Geller", "Joey Tribbiani"]
    
    domains = ["gmail.com", "techcorp.io", "dundermifflin.com", "enterprise.org", "outlook.com"]
    bad_emails = ["invalid_email_at_test", "pam.b@gmail..com", "dwight_dunder..com", "m.scott@office"]
    phones = ["1234567890", "+1 (987) 654-3210", "555-0199", "15550201122", "123 456 7890", "9876543210"]
    dates = ["2023-01-15", "02/20/2024", "2024/12/05", "15-08-2025", "2026-01-10", "11/05/2023", "2024-09-30", "2025-04-18"]
    regions = ["North America", "Europe", "Asia-Pacific", "Latin America", "Middle East"]
    categories = ["Enterprise Software", "Hardware", "Cloud Services", "Consulting", "Support Plans"]
    statuses = [" COMPLETED ", "pending", "CANCELLED", " Completed ", "Shipped", "Processing"]
    reps = ["Alice Smith", "Bob Jones", "Charlie Brown", "Diana Prince", "Evan Wright"]

    rows = []
    for i in range(1, n_rows + 1):
        if i in [20, 100, 250, 500, 750, 950]:
            rows.append([np.nan] * 15)
            continue

        trans_id = f" TR-2026-{(i % 800) + 1000:04d} "
        cust_name = names[i % len(names)]
        
        if i % 11 == 0:
            email = bad_emails[i % len(bad_emails)]
        else:
            clean_n = cust_name.lower().replace(" ", ".")
            email = f"{clean_n}@{domains[i % len(domains)]}"

        phone = phones[i % len(phones)]
        reg_date = dates[i % len(dates)]
        region = regions[i % len(regions)]
        category = categories[i % len(categories)]
        status = statuses[i % len(statuses)]
        sales_rep = reps[i % len(reps)]
        
        units = int((i * 9 % 120) + 5)
        unit_price = round(float((i * 17 % 450) + 49.99), 2)
        total_revenue = round(units * unit_price, 2)
        operational_cost = round(total_revenue * float(((i % 5) + 2) * 0.11), 2)
        net_profit = round(total_revenue - operational_cost, 2)
        discount_rate = round(float((i % 15) * 0.01), 2)
        csat = round(float((i % 5) + 1.0), 1)

        rows.append([
            trans_id, cust_name, email, phone, reg_date, region, 
            category, status, sales_rep, units, unit_price, 
            total_revenue, operational_cost, net_profit, csat
        ])

    cols = [
        " Transaction ID ", "Customer Name", "Email Address", "Phone Number", 
        "Transaction Date", "Region", "Category", "Status", "Sales Rep", 
        "Units Sold", "Unit Price", "Total Revenue", "Operational Cost", "Net Profit", "CSAT Score"
    ]
    df = pd.DataFrame(rows, columns=cols)

    df = pd.concat([df, df.iloc[[10, 45, 120, 300, 550, 800]]], ignore_index=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Global_Sales_Master')
    return output.getvalue()

@st.cache_data
def generate_large_csv_sample():
    """Generates a high-capacity CSV customer dataset designed for matching and joining."""
    rows = []
    regions = ["North America", "Europe", "Asia-Pacific", "Latin America", "Middle East"]
    tiers = ["Enterprise", "Mid-Market", "SMB", "Government", "Startup"]
    
    for i in range(1000, 2200):
        cust_id = f"CUST-{i:04d}"
        comp_name = f"Global Tech Corp {i}"
        reg = regions[i % len(regions)]
        tier = tiers[i % len(tiers)]  # Fixed UnboundLocalError bug
        credit_limit = float((i * 2500) % 250000 + 15000)
        account_balance = round(credit_limit * 0.35 + (i * 12 % 5000), 2)
        churn_risk_score = round(float((i % 10) * 0.09), 2)
        active_status = "Active" if i % 6 != 0 else "Inactive"
        
        rows.append({
            "CustomerID": cust_id,
            "CompanyName": comp_name,
            "Region": reg,
            "AccountTier": tier,
            "CreditLimit": credit_limit,
            "AccountBalance": account_balance,
            "ChurnRiskScore": churn_risk_score,
            "AccountStatus": active_status
        })
        
    df = pd.DataFrame(rows)
    return df.to_csv(index=False).encode('utf-8')

# ==========================================
# HELPER DATA CLEANING LOGIC
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

def execute_data_cleaning(df, opts):
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
# EXECUTIVE SIDEBAR CONTROL PANEL
# ==========================================
st.sidebar.markdown('<div class="sidebar-header">⚡ Excel Toolkit Pro</div>', unsafe_allow_html=True)
st.sidebar.markdown('<div class="sidebar-badge">v2.5 Enterprise Edition</div>', unsafe_allow_html=True)

st.sidebar.markdown("---")
st.sidebar.markdown("### 🔑 License Activation")
input_key = st.sidebar.text_input("Lemon Squeezy License Key", type="password", help="Enter your product key to unlock custom file processing.")

if st.sidebar.button("Activate License Key", use_container_width=True):
    is_valid, msg = verify_lemon_squeezy_license(input_key)
    st.session_state.is_subscribed = is_valid
    st.session_state.license_msg = msg

status_color = "#10B981" if st.session_state.is_subscribed else "#F59E0B"
status_icon = "🟢" if st.session_state.is_subscribed else "🔒"

st.sidebar.markdown(f"""
<div class="sidebar-status-box">
    <div class="sidebar-status-title">Current Access Status</div>
    <div class="sidebar-status-val" style="color: {status_color} !important;">{status_icon} {st.session_state.license_msg}</div>
</div>
""", unsafe_allow_html=True)

if not st.session_state.is_subscribed:
    st.sidebar.markdown(
        f'<a href="{LEMON_SQUEEZY_CHECKOUT_URL}" target="_blank" style="text-decoration:none;">'
        f'<button style="width:100%; background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%); color:#FFF !important; padding:10px; border-radius:8px; font-weight:600; border:none; cursor:pointer; margin-bottom:15px;">'
        f'🛒 Upgrade to Pro Version</button></a>',
        unsafe_allow_html=True
    )

st.sidebar.markdown("---")

module = st.sidebar.radio(
    "Navigation Modules",
    [
        "🧹 Data Cleaning Studio",
        "🗂️ File & Batch Operations",
        "🔍 Dataset Matching & Compare",
        "⚡ Power Tools & Conversion",
        "📊 Interactive Dashboard Studio"
    ]
)

st.sidebar.markdown("---")
st.sidebar.caption("© 2026 Enterprise Analytics Studio")

# ==========================================
# HERO & FREE SAMPLE DOWNLOAD SECTION
# ==========================================
st.markdown("""
<div class="hero-container">
    <div class="hero-title">Excel Automation & Analytics Studio Pro</div>
    <p class="hero-subtitle">High-performance spreadsheet scrubbing, multi-file batch operations, reconciliation, and automated dashboard generation.</p>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="free-banner">
    <div style="font-weight:700; font-size:1.05rem; color:#1E3A8A; margin-bottom:0.3rem;">⚡ Test Drive Free Sample Datasets</div>
    <div style="font-size:0.9rem; color:#3B82F6; margin-bottom:0.8rem;">
        No active license? Download our new 1,000+ record sample datasets below to test all tools, cleanings, and visualization studios for free. Custom file processing requires a Pro license key.
    </div>
</div>
""", unsafe_allow_html=True)

btn1, btn2, _ = st.columns([1.2, 1.2, 1.6])

btn1.download_button(
    label="📥 Excel Data (1,000+ Rows)",
    data=generate_large_excel_sample(),
    file_name="Enterprise_Global_Sales_2026.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)

btn2.download_button(
    label="📥 CSV Data (1,200 Rows)",
    data=generate_large_csv_sample(),
    file_name="Enterprise_Customer_Master_2026.csv",
    mime="text/csv",
    use_container_width=True
)

st.markdown("<br>", unsafe_allow_html=True)

def show_subscription_required_warning(filename=""):
    st.error(
        f"🔒 **Pro License Required for Custom Uploads**\n\n"
        f"Processing custom files like `{filename}` requires an active Pro license. "
        f"Please enter your Lemon Squeezy license key in the sidebar, or test the tool for free using the sample datasets downloaded above."
    )

# ==========================================
# MODULE 1: DATA CLEANING STUDIO
# ==========================================
if module == "🧹 Data Cleaning Studio":
    st.markdown("### 🧹 Data Cleaning Studio")
    st.caption("Automate data scrubbing, standardizations, deduplication, and pattern cleanings.")

    uploaded_file = st.file_uploader("Upload File (Excel or CSV)", type=["xlsx", "xls", "csv"])

    if uploaded_file:
        file_ext = uploaded_file.name.split('.')[-1].lower()
        df = pd.read_csv(uploaded_file) if file_ext == 'csv' else pd.read_excel(uploaded_file)

        st.markdown("#### Raw Dataset Preview")
        st.dataframe(df.head(6), use_container_width=True)

        col1, col2, col3 = st.columns(3)

        with col1:
            st.markdown("##### Basic Scrubbing")
            rem_dupes = st.checkbox("Remove Duplicate Records", value=True)
            rem_blanks = st.checkbox("Remove Entirely Blank Rows", value=True)
            trim_sp = st.checkbox("Trim Whitespace Across All Cells", value=True)
            std_cols = st.checkbox("Convert Column Names to snake_case", value=False)

        with col2:
            st.markdown("##### Formatting & Formatting")
            cap_type = st.selectbox("Text Capitalization Standard", ["None", "UPPERCASE", "lowercase", "Title Case"])
            clean_em = st.checkbox("Format & Validate Email Addresses")
            email_cols = st.multiselect("Target Email Columns", df.columns) if clean_em else []

            clean_ph = st.checkbox("Standardize Phone Numbers (+1 US Format)")
            phone_cols = st.multiselect("Target Phone Columns", df.columns) if clean_ph else []

        with col3:
            st.markdown("##### Dates & Subset Rules")
            std_dt = st.checkbox("Standardize Date Column Formats")
            date_cols = st.multiselect("Target Date Columns", df.columns) if std_dt else []
            dt_fmt = st.selectbox("Output Date Format", ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"]) if std_dt else "%Y-%m-%d"

            dupe_sub = st.multiselect("Deduplicate Based On Unique Key Column(s)", df.columns) if rem_dupes else []

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("⚡ Execute Cleaning Pipeline", use_container_width=True):
            if not is_allowed_file(uploaded_file):
                show_subscription_required_warning(uploaded_file.name)
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

                cleaned_df, report = execute_data_cleaning(df, options)
                st.success("Cleaning Pipeline Executed Successfully!")

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Original Rows", f"{report['initial_rows']:,}")
                m2.metric("Cleaned Rows", f"{report['final_rows']:,}")
                m3.metric("Duplicates Removed", report["duplicates_removed"])
                m4.metric("Blank Rows Removed", report["blanks_removed"])

                st.markdown("#### Cleaned Dataset Preview")
                st.dataframe(cleaned_df.head(10), use_container_width=True)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    cleaned_df.to_excel(writer, index=False, sheet_name='Cleaned_Data')

                st.download_button(
                    label="📥 Download Cleaned Excel Dataset",
                    data=output.getvalue(),
                    file_name=f"Cleaned_{uploaded_file.name.split('.')[0]}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

# ==========================================
# MODULE 2: FILE & BATCH OPERATIONS
# ==========================================
elif module == "🗂️ File & Batch Operations":
    st.markdown("### 🗂️ File & Batch Operations")
    st.caption("Merge, split, extract, and bulk transform spreadsheet workbooks.")

    action = st.selectbox(
        "Select Operation Tool",
        [
            "Merge Multiple Excel/CSV Files",
            "Batch Text Search & Replace Across Files",
            "Split Workbook by Column Category",
            "Split Workbook by Transaction Year",
            "Batch CSV → Excel Converter",
            "Column Extraction Utility"
        ]
    )

    if action == "Merge Multiple Excel/CSV Files":
        files = st.file_uploader("Upload Workbooks to Combine", type=["csv", "xlsx", "xls"], accept_multiple_files=True)
        if files:
            if st.button("Combine Files Into Master Workbook"):
                if not is_allowed_file(files):
                    show_subscription_required_warning("Batch Upload Group")
                else:
                    dfs = []
                    for f in files:
                        ext = f.name.split('.')[-1].lower()
                        df_temp = pd.read_csv(f) if ext == 'csv' else pd.read_excel(f)
                        df_temp['Source_Workbook'] = f.name
                        dfs.append(df_temp)
                    merged_df = pd.concat(dfs, ignore_index=True)
                    st.success(f"Successfully merged {len(files)} files into {len(merged_df):,} combined records!")

                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        merged_df.to_excel(writer, index=False, sheet_name='Merged_Master')
                    st.download_button("📥 Download Merged Dataset", data=output.getvalue(), file_name="Merged_Master_Workbook.xlsx", use_container_width=True)

    elif action == "Batch Text Search & Replace Across Files":
        files = st.file_uploader("Upload Files for Batch Replacement", type=["csv", "xlsx"], accept_multiple_files=True)
        if files:
            col_search, col_replace = st.columns(2)
            search_str = col_search.text_input("Text / String to Find")
            replace_str = col_replace.text_input("Replacement Value")

            if search_str and st.button("Run Batch Replacement"):
                if not is_allowed_file(files):
                    show_subscription_required_warning("Batch Files Group")
                else:
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                        for f in files:
                            ext = f.name.split('.')[-1].lower()
                            df_temp = pd.read_csv(f) if ext == 'csv' else pd.read_excel(f)
                            
                            for col in df_temp.select_dtypes(include=['object', 'string']).columns:
                                df_temp[col] = df_temp[col].astype(str).str.replace(search_str, replace_str, regex=False)

                            out_b = io.BytesIO()
                            with pd.ExcelWriter(out_b, engine='openpyxl') as writer:
                                df_temp.to_excel(writer, index=False, sheet_name='Updated_Data')
                            zip_file.writestr(f"Updated_{f.name.rsplit('.', 1)[0]}.xlsx", out_b.getvalue())

                    st.success("Batch replacement completed across all files!")
                    st.download_button("📥 Download Updated Files (.zip)", data=zip_buffer.getvalue(), file_name="Batch_Updated_Files.zip", use_container_width=True)

    elif action == "Split Workbook by Column Category":
        uploaded_file = st.file_uploader("Upload File to Split", type=["xlsx", "csv"])
        if uploaded_file:
            df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
            split_col = st.selectbox("Select Splitting Key Column", df.columns)

            if st.button("Generate Categorical Split Archive"):
                if not is_allowed_file(uploaded_file):
                    show_subscription_required_warning(uploaded_file.name)
                else:
                    unique_vals = df[split_col].dropna().unique()
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                        for val in unique_vals:
                            sub_df = df[df[split_col] == val]
                            out_b = io.BytesIO()
                            with pd.ExcelWriter(out_b, engine='openpyxl') as writer:
                                sub_df.to_excel(writer, index=False, sheet_name=str(val)[:30])
                            clean_val_str = re.sub(r'[^\w\-_\. ]', '_', str(val))
                            zip_file.writestr(f"Split_{split_col}_{clean_val_str}.xlsx", out_b.getvalue())

                    st.success(f"Split complete! {len(unique_vals)} standalone workbooks generated.")
                    st.download_button("📥 Download Split Archive (.zip)", data=zip_buffer.getvalue(), file_name="Categorical_Split_Archive.zip", use_container_width=True)

    elif action == "Split Workbook by Transaction Year":
        uploaded_file = st.file_uploader("Upload File", type=["xlsx", "csv"])
        if uploaded_file:
            df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
            date_col = st.selectbox("Select Date Column", df.columns)

            if st.button("Execute Split by Year"):
                if not is_allowed_file(uploaded_file):
                    show_subscription_required_warning(uploaded_file.name)
                else:
                    df['__Temp_Year'] = pd.to_datetime(df[date_col], errors='coerce').dt.year
                    years = df['__Temp_Year'].dropna().unique()

                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                        for y in years:
                            sub_df = df[df['__Temp_Year'] == y].drop(columns=['__Temp_Year'])
                            out_b = io.BytesIO()
                            with pd.ExcelWriter(out_b, engine='openpyxl') as writer:
                                sub_df.to_excel(writer, index=False, sheet_name=str(int(y)))
                            zip_file.writestr(f"Data_Year_{int(y)}.xlsx", out_b.getvalue())

                    st.success("Successfully generated annual files!")
                    st.download_button("📥 Download Year Split Archive (.zip)", data=zip_buffer.getvalue(), file_name="Year_Split_Files.zip", use_container_width=True)

    elif action == "Batch CSV → Excel Converter":
        files = st.file_uploader("Upload CSV Files", type=["csv"], accept_multiple_files=True)
        if files and st.button("Convert All CSVs to Formatted Excel"):
            if not is_allowed_file(files):
                show_subscription_required_warning("CSV Conversion Group")
            else:
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    for f in files:
                        df_c = pd.read_csv(f)
                        out_b = io.BytesIO()
                        with pd.ExcelWriter(out_b, engine='openpyxl') as writer:
                            df_c.to_excel(writer, index=False, sheet_name='Sheet1')
                        base_name = f.name.rsplit('.', 1)[0]
                        zip_file.writestr(f"{base_name}.xlsx", out_b.getvalue())
                st.download_button("📥 Download Converted Excel Workbooks (.zip)", data=zip_buffer.getvalue(), file_name="Batch_Converted_Excel.zip", use_container_width=True)

    elif action == "Column Extraction Utility":
        uploaded_file = st.file_uploader("Upload File", type=["xlsx", "csv"])
        if uploaded_file:
            df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
            selected_cols = st.multiselect("Select Target Columns to Isolate", df.columns)
            if selected_cols and st.button("Extract & Download Selected Fields"):
                if not is_allowed_file(uploaded_file):
                    show_subscription_required_warning(uploaded_file.name)
                else:
                    extracted_df = df[selected_cols]
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        extracted_df.to_excel(writer, index=False)
                    st.download_button("📥 Download Extracted Fields File", data=output.getvalue(), file_name="Extracted_Columns.xlsx", use_container_width=True)

# ==========================================
# MODULE 3: DATASET MATCHING & COMPARE
# ==========================================
elif module == "🔍 Dataset Matching & Compare":
    st.markdown("### 🔍 Dataset Matching & Sheet Reconciliation")
    st.caption("Perform multi-table joins, cross-workbook VLOOKUP reconciliation, and row-level diff tracking.")

    mode = st.radio("Select Reconciliation Tool", ["Match Two Datasets (VLOOKUP / Join Engine)", "Compare Two Excel Files Cell-by-Cell"])

    if mode == "Match Two Datasets (VLOOKUP / Join Engine)":
        c1, c2 = st.columns(2)
        with c1:
            f1 = st.file_uploader("Upload Primary File (Dataset A)", type=["xlsx", "csv"], key="m_f1")
        with c2:
            f2 = st.file_uploader("Upload Reference File (Dataset B)", type=["xlsx", "csv"], key="m_f2")

        if f1 and f2:
            df1 = pd.read_csv(f1) if f1.name.endswith('.csv') else pd.read_excel(f1)
            df2 = pd.read_csv(f2) if f2.name.endswith('.csv') else pd.read_excel(f2)

            col_a, col_b = st.columns(2)
            key1 = col_a.selectbox("Primary Match Key (Dataset A)", df1.columns)
            key2 = col_b.selectbox("Reference Match Key (Dataset B)", df2.columns)

            join_type = st.selectbox("Join Algorithm", ["Inner Join (Matched Records Only)", "Left Join (Retain All Dataset A)", "Find Unmatched Keys in Dataset A"])

            if st.button("Execute Dataset Match", use_container_width=True):
                if not (is_allowed_file(f1) and is_allowed_file(f2)):
                    show_subscription_required_warning(f"{f1.name} / {f2.name}")
                else:
                    df1_match = df1.copy()
                    df2_match = df2.copy()
                    df1_match[key1] = df1_match[key1].astype(str).str.strip()
                    df2_match[key2] = df2_match[key2].astype(str).str.strip()

                    if join_type == "Inner Join (Matched Records Only)":
                        res = pd.merge(df1_match, df2_match, left_on=key1, right_on=key2, how='inner', suffixes=('_A', '_B'))
                    elif join_type == "Left Join (Retain All Dataset A)":
                        res = pd.merge(df1_match, df2_match, left_on=key1, right_on=key2, how='left', suffixes=('_A', '_B'))
                    else:
                        res = df1_match[~df1_match[key1].isin(df2_match[key2])]

                    st.success(f"Matched {len(res):,} records!")
                    st.dataframe(res.head(10), use_container_width=True)

                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        res.to_excel(writer, index=False)
                    st.download_button("📥 Download Reconciled Dataset", data=output.getvalue(), file_name="Dataset_Match_Report.xlsx", use_container_width=True)

    elif mode == "Compare Two Excel Files Cell-by-Cell":
        c1, c2 = st.columns(2)
        with c1:
            f1 = st.file_uploader("Original File A", type=["xlsx", "csv"], key="c_f1")
        with c2:
            f2 = st.file_uploader("Modified File B", type=["xlsx", "csv"], key="c_f2")

        if f1 and f2:
            df1 = pd.read_csv(f1) if f1.name.endswith('.csv') else pd.read_excel(f1)
            df2 = pd.read_csv(f2) if f2.name.endswith('.csv') else pd.read_excel(f2)

            if st.button("Run Cell-by-Cell Comparison", use_container_width=True):
                if not (is_allowed_file(f1) and is_allowed_file(f2)):
                    show_subscription_required_warning(f"{f1.name} / {f2.name}")
                else:
                    common_cols = list(set(df1.columns).intersection(set(df2.columns)))
                    df1_c = df1[common_cols]
                    df2_c = df2[common_cols]

                    diff = df1_c.compare(df2_c)
                    if diff.empty:
                        st.success("✅ Datasets are 100% identical across all overlapping fields!")
                    else:
                        st.warning(f"Detected {len(diff):,} row differences across common attributes.")
                        st.dataframe(diff, use_container_width=True)

                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            diff.to_excel(writer)
                        st.download_button("📥 Download Row Difference Log", data=output.getvalue(), file_name="Cell_Difference_Log.xlsx", use_container_width=True)

# ==========================================
# MODULE 4: POWER TOOLS & CONVERSION
# ==========================================
elif module == "⚡ Power Tools & Conversion":
    st.markdown("### ⚡ Power Tools & Advanced Operations")
    st.caption("Execute complex text filters, multi-column sorting, and bulk tab renames.")

    p_action = st.selectbox("Select Utility Tool", ["Filter & Multi-Column Sorting", "Batch Rename Excel Worksheets"])

    if p_action == "Filter & Multi-Column Sorting":
        uploaded_file = st.file_uploader("Upload File", type=["xlsx", "csv"])
        if uploaded_file:
            df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)

            c1, c2 = st.columns(2)
            filter_col = c1.selectbox("Filter Column Target", ["None"] + list(df.columns))
            sort_col = c2.selectbox("Sort Priority Column", ["None"] + list(df.columns))

            filtered_df = df.copy()

            if filter_col != "None":
                val = st.text_input(f"Substring Match for '{filter_col}'")
                if val:
                    filtered_df = filtered_df[filtered_df[filter_col].astype(str).str.contains(val, case=False, na=False)]

            if sort_col != "None":
                ascending = st.checkbox("Sort Ascending", value=True)
                filtered_df = filtered_df.sort_values(by=sort_col, ascending=ascending)

            st.dataframe(filtered_df.head(10), use_container_width=True)

            if st.button("Export Processed Data"):
                if not is_allowed_file(uploaded_file):
                    show_subscription_required_warning(uploaded_file.name)
                else:
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        filtered_df.to_excel(writer, index=False)
                    st.download_button("📥 Download Filtered Dataset", data=output.getvalue(), file_name="Filtered_Sorted_Export.xlsx", use_container_width=True)

    elif p_action == "Batch Rename Excel Worksheets":
        uploaded_file = st.file_uploader("Upload Multi-Sheet Excel File", type=["xlsx"])
        if uploaded_file:
            xl = pd.ExcelFile(uploaded_file)
            sheet_names = xl.sheet_names
            st.info(f"Detected {len(sheet_names)} worksheets: {', '.join(sheet_names)}")

            prefix = st.text_input("Add Prefix to All Tab Names", value="")
            suffix = st.text_input("Add Suffix to All Tab Names", value="")

            if st.button("Execute Tab Renaming"):
                if not is_allowed_file(uploaded_file):
                    show_subscription_required_warning(uploaded_file.name)
                else:
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        for s in sheet_names:
                            df_s = pd.read_excel(uploaded_file, sheet_name=s)
                            new_name = f"{prefix}{s}{suffix}"[:31]
                            df_s.to_excel(writer, sheet_name=new_name, index=False)
                    st.success("Worksheet tab names updated!")
                    st.download_button("📥 Download Renamed Workbook", data=output.getvalue(), file_name="Renamed_Tabs_Workbook.xlsx", use_container_width=True)

# ==========================================
# MODULE 5: INTERACTIVE DASHBOARD STUDIO & ANALYZER
# ==========================================
elif module == "📊 Interactive Dashboard Studio":
    st.markdown("### 📊 Interactive Dashboard Studio & Advanced Analyzer")
    st.caption("Generate executive metrics, comprehensive statistical diagnostics, and Plotly visualizations.")

    uploaded_file = st.file_uploader("Upload Excel or CSV File for Analytics", type=["xlsx", "csv"], key="dash_file")

    if uploaded_file:
        df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
        
        if not is_allowed_file(uploaded_file):
            show_subscription_required_warning(uploaded_file.name)
        else:
            num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
            cat_cols = df.select_dtypes(include=['object', 'string', 'category']).columns.tolist()

            st.markdown("#### ⚡ Executive Summary Metrics")
            k1, k2, k3, k4 = st.columns(4)

            k1.markdown(f"""
            <div class="kpi-card-container">
                <div class="kpi-title">Total Rows</div>
                <div class="kpi-val">{len(df):,}</div>
            </div>
            """, unsafe_allow_html=True)

            k2.markdown(f"""
            <div class="kpi-card-container">
                <div class="kpi-title">Attributes / Fields</div>
                <div class="kpi-val">{len(df.columns)}</div>
            </div>
            """, unsafe_allow_html=True)

            if num_cols:
                primary_metric = st.selectbox("Primary KPI Target Metric", num_cols, index=0)
                tot_val = df[primary_metric].sum()
                avg_val = df[primary_metric].mean()

                k3.markdown(f"""
                <div class="kpi-card-container">
                    <div class="kpi-title">Total {primary_metric}</div>
                    <div class="kpi-val">{tot_val:,.2f}</div>
                </div>
                """, unsafe_allow_html=True)

                k4.markdown(f"""
                <div class="kpi-card-container">
                    <div class="kpi-title">Average {primary_metric}</div>
                    <div class="kpi-val">{avg_val:,.2f}</div>
                </div>
                """, unsafe_allow_html=True)
            else:
                k3.markdown(f"""
                <div class="kpi-card-container">
                    <div class="kpi-title">Categorical Fields</div>
                    <div class="kpi-val">{len(cat_cols)}</div>
                </div>
                """, unsafe_allow_html=True)
                
                k4.markdown(f"""
                <div class="kpi-card-container">
                    <div class="kpi-title">Numeric Fields</div>
                    <div class="kpi-val">0</div>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

            # Statistical Profiler
            with st.expander("📈 Advanced Statistical Diagnostic Summary"):
                if num_cols:
                    stat_df = df[num_cols].describe().T
                    stat_df['median'] = df[num_cols].median()
                    stat_df['skewness'] = df[num_cols].skew()
                    st.dataframe(stat_df[['count', 'mean', 'median', 'std', 'min', 'max', 'skewness']].style.format("{:,.2f}"), use_container_width=True)
                else:
                    st.info("No numeric features available for statistical analysis.")

            st.markdown("---")
            st.markdown("#### 📈 Visual Analytics Builder")

            dash_type = st.selectbox(
                "Select Visualization Engine",
                [
                    "Bar Chart (Categorical Aggregation)",
                    "Line Chart (Trend & Time Series)",
                    "Scatter Plot (Correlation Analysis)",
                    "Pie / Donut Chart (Proportion Share)",
                    "Histogram (Frequency Distribution)",
                    "Heatmap (Correlation Matrix)"
                ]
            )

            c1, c2 = st.columns(2)

            if dash_type == "Bar Chart (Categorical Aggregation)":
                if cat_cols and num_cols:
                    x_col = c1.selectbox("Category Dimension (X-Axis)", cat_cols)
                    y_col = c2.selectbox("Metric Target (Y-Axis)", num_cols)
                    agg_func = st.selectbox("Aggregation Rule", ["Sum", "Mean", "Count"])
                    
                    if agg_func == "Sum":
                        grouped = df.groupby(x_col, as_index=False)[y_col].sum()
                    elif agg_func == "Mean":
                        grouped = df.groupby(x_col, as_index=False)[y_col].mean()
                    else:
                        grouped = df.groupby(x_col, as_index=False)[y_col].count()

                    fig = px.bar(
                        grouped, x=x_col, y=y_col,
                        title=f"<b>{agg_func} of {y_col} by {x_col}</b>",
                        template="plotly_white",
                        color_discrete_sequence=["#2563EB"]
                    )
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.warning("Bar charts require at least 1 categorical column and 1 numeric column.")

            elif dash_type == "Line Chart (Trend & Time Series)":
                if num_cols:
                    x_col = c1.selectbox("Time or Axis Dimension", df.columns)
                    y_col = c2.selectbox("Metric Target", num_cols)
                    fig = px.line(
                        df, x=x_col, y=y_col,
                        title=f"<b>{y_col} Over {x_col}</b>",
                        template="plotly_white",
                        markers=True,
                        color_discrete_sequence=["#0EA5E9"]
                    )
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.warning("Line charts require at least 1 numeric metric column.")

            elif dash_type == "Scatter Plot (Correlation Analysis)":
                if len(num_cols) >= 2:
                    x_col = c1.selectbox("X-Axis Metric", num_cols, index=0)
                    y_col = c2.selectbox("Y-Axis Metric", num_cols, index=min(1, len(num_cols)-1))
                    color_col = st.selectbox("Group / Color Dimension (Optional)", ["None"] + cat_cols)
                    
                    color_param = None if color_col == "None" else color_col
                    fig = px.scatter(
                        df, x=x_col, y=y_col, color=color_param,
                        title=f"<b>Correlation: {x_col} vs {y_col}</b>",
                        template="plotly_white"
                    )
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.warning("Scatter plots require at least 2 numeric columns.")

            elif dash_type == "Pie / Donut Chart (Proportion Share)":
                if cat_cols:
                    names_col = c1.selectbox("Category Dimension", cat_cols)
                    values_col = c2.selectbox("Value Metric (Optional)", ["Record Count"] + num_cols)
                    hole_val = 0.45 if st.checkbox("Render as Modern Donut Chart", value=True) else 0.0
                    
                    if values_col == "Record Count":
                        fig = px.pie(df, names=names_col, title=f"<b>Distribution by {names_col}</b>", hole=hole_val, template="plotly_white")
                    else:
                        fig = px.pie(df, names=names_col, values=values_col, title=f"<b>{values_col} Share by {names_col}</b>", hole=hole_val, template="plotly_white")
                    
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.warning("Pie charts require at least 1 categorical column.")

            elif dash_type == "Histogram (Frequency Distribution)":
                if num_cols:
                    num_col = c1.selectbox("Target Numeric Field", num_cols)
                    bins = c2.slider("Histogram Bins", min_value=5, max_value=60, value=20)
                    fig = px.histogram(
                        df, x=num_col, nbins=bins,
                        title=f"<b>Frequency Distribution of {num_col}</b>",
                        template="plotly_white",
                        color_discrete_sequence=["#3B82F6"]
                    )
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.warning("Histograms require at least 1 numeric column.")

            elif dash_type == "Heatmap (Correlation Matrix)":
                if len(num_cols) >= 2:
                    corr = df[num_cols].corr()
                    fig = px.imshow(
                        corr, text_auto=".2f", aspect="auto",
                        title="<b>Numeric Feature Correlation Matrix</b>",
                        color_continuous_scale="Blues",
                        template="plotly_white"
                    )
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.warning("Correlation Heatmaps require at least 2 numeric columns.")
