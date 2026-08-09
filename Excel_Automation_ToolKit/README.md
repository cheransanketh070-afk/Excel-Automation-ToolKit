# 🛠️ Excel Automation Toolkit Pro

**Excel Automation Toolkit Pro** is a high-performance web application built with Python, Streamlit, and Plotly. Designed for data analysts, marketers, operations teams, and finance professionals, it automates tedious spreadsheet scrubbing, batch transformations, dataset reconciliation, and executive chart generation—without requiring complex Excel formulas, VLOOKUPs, or VBA macros.

The application features seamless monetization and access control powered by **Lemon Squeezy**, including license key validation, direct checkout routing, and built-in sample file testing.

---

## 🌟 Key Features & Functional Modules

### 1. 🧹 Data Cleaning Studio
* **Automated Data Scrubbing**: Eliminate blank rows, trim leading/trailing whitespace, and standardize text across entire worksheets.
* **Smart Field Scrubbing**: Automatically extract and reformat phone numbers into standard US formats (`+1 (XXX) XXX-XXXX`) and clean invalid email addresses using regex pattern matching.
* **Standardization**: Convert column headers into clean `snake_case` naming conventions and unify date formats (`YYYY-MM-DD`, `MM/DD/YYYY`, etc.).
* **Deduplication**: Identify and remove duplicate records globally or based on targeted primary key columns.

### 2. 🗂️ File & Batch Operations
* **Merge Workbooks**: Concatenate multiple Excel (`.xlsx`) or CSV files into a single unified workbook with automatic source-file tracking.
* **Split Utilities**: Automatically split large datasets into individual sheets or standalone workbooks by column category or transaction year.
* **Batch Conversion**: Convert bulk CSV files into formatted Excel (`.xlsx`) workbooks packaged in a single ZIP archive.
* **Column Extractor**: Select and isolate specific data columns across datasets into a clean export file.

### 3. 🔍 Dataset Matching & Sheet Comparison
* **VLOOKUP/XLOOKUP Replacement**: Perform inner joins, left outer joins, or identify missing records between two separate datasets without formula overhead.
* **Workbook Differencing**: Compare two versions of an Excel sheet cell-by-cell to highlight, track, and export modified values into a detailed difference log.

### 4. ⚡ Power Tools & Advanced Operations
* **Filter & Search**: Query and slice complex workbooks using text containment filters and dynamic multi-column sorting.
* **Batch Sheet Renaming**: Apply custom prefixes or suffixes across all worksheets within a multi-sheet workbook at once.

### 5. 📊 Interactive Dashboard Studio *(New)*
* **Executive KPI Cards**: Instantly compute and display key summary metrics including record counts, column dimensions, totals, and averages.
* **6 Modern Visualizations**: Build interactive Plotly charts directly from uploaded datasets:
  * Bar Charts (Category Aggregation)
  * Line Charts (Trend & Time Series)
  * Scatter Plots (Correlation Analysis with multi-variable color grouping)
  * Pie & Donut Charts (Distribution percentages)
  * Histograms (Frequency distributions)
  * Heatmaps (Numeric correlation matrices)

---

## 🔐 Licensing & Subscription Controls

The app includes direct integration with **Lemon Squeezy** for software monetization:

* **Live Key Validation**: Interacts with the Lemon Squeezy API (`v1/licenses/activate`) to verify customer license status.
* **Embedded Checkout**: A native sidebar checkout button directs unsubscribed users to your active Lemon Squeezy store.
* **Developer Bypass Mode**: Features an embedded test key (`TEST-GUM-7172504D6E9D`) for offline debugging and administrative overrides.
* **Free Interactive Sample Mode**:
  * Unsubscribed users can download pre-packaged test datasets (`01_Messy_Customer_Data.xlsx` and `02_Dataset_A_Customers.csv`) directly from the top banner to test all 5 modules risk-free.
  * Uploading or processing custom proprietary files without an active license triggers a security gate warning.

---

## 📁 Repository Structure

```text
├── app.py                         # Complete Streamlit application code
├── requirements.txt               # Python package dependencies
├── 01_Messy_Customer_Data.xlsx    # Sample Excel file (Auto-generated fallback)
├── 02_Dataset_A_Customers.csv     # Sample CSV file (Auto-generated fallback)
└── README.md                      # Complete project documentation