# ⚡ Excel Automation & Analytics Studio Pro

An enterprise-grade Streamlit suite for spreadsheet data cleaning, multi-file batch operations, VLOOKUP reconciliation, and interactive Plotly analytics.

---

## ✨ Features & Capabilities

### 🧹 1. Data Cleaning Studio
* **Automated Data Scrubbing**: Eliminate blank rows, duplicate entries, and unwanted whitespace.
* **Smart Formatting**: Standardize text case (UPPERCASE, lowercase, Title Case) and column names to `snake_case`.
* **Pattern Validation**: Format and validate email addresses and US phone numbers (`+1 (XXX) XXX-XXXX`).
* **Date Standardization**: Unify inconsistent date strings across columns into standard formats (`YYYY-MM-DD`, `MM/DD/YYYY`).

### 🗂️ 2. File & Batch Operations
* **Merge Workbooks**: Concatenate multiple CSV/Excel files into a unified master dataset with source tracking.
* **Batch Search & Replace**: Execute bulk string/value replacements across multi-file workbook archives.
* **Categorical & Date Splitting**: Split massive sheets into separate workbooks based on unique column categories or transaction years.
* **Format Conversion & Field Extraction**: Convert batch CSV files into formatted Excel workbooks or isolate specific columns into standalone files.

### 🔍 3. Dataset Matching & Compare
* **Multi-Table Joins**: Perform cross-workbook joins (Inner, Left, or Unmatched record isolation).
* **Cell-by-Cell Reconciliation**: Compare two versions of a sheet to highlight precise row-and-column level differences.

### ⚡ 4. Power Tools & Custom Rules
* **Multi-Column Filtering & Sorting**: Filter datasets by substring patterns and re-order records dynamically.
* **Batch Worksheet Renaming**: Update tab titles across complex workbooks using custom prefixes and suffixes.

### 📊 5. Interactive Dashboard Studio & Analyzer
* **KPI Metrics**: Modern glassmorphism stat cards for high-level aggregate overviews.
* **Statistical Diagnostics**: Automated profiling including mean, median, standard deviation, skewness, min/max, and quantiles.
* **Plotly Visualizations**: Interactive Bar Charts, Time Series Line Graphs, Scatter Plots, Donut/Pie Charts, Histograms, and Correlation Heatmaps.

---

## 🔒 Licensing & Free Tier Access

* **Free Tier / Sample Testing**: Free users can download and test all tools using built-in enterprise sample datasets (`Enterprise_Global_Sales_2026.xlsx` & `Enterprise_Customer_Master_2026.csv`).
* **Pro Tier**: Uploading and processing custom user files is secured via Lemon Squeezy license key verification.

---

## 🚀 Quickstart Guide

### 1. Prerequisites
Ensure you have Python 3.9+ installed on your system.

### 2. Installation
Clone the repository and install the dependencies:

```bash
# Install required libraries
pip install -r requirements.txt
