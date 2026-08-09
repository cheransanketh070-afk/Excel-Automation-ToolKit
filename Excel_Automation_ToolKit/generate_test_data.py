import os
import random
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

# Create output folder
OUTPUT_DIR = "sample_test_datasets"
os.makedirs(OUTPUT_DIR, exist_ok=True)

print("🚀 Generating sample test datasets for Excel Automation Toolkit...\n")

# ==========================================
# 1. MESSY DATASET FOR DATA CLEANING STUDIO
# ==========================================
def generate_messy_cleaning_data():
    np.random.seed(42)
    random.seed(42)

    raw_names = [" john DOE ", "SARAH connor", "michael scott ", "Pam Beesly", " Dwight Schrute ", "jim HALPERT"]
    raw_emails = ["JOHN.DOE@GMAIL.COM", "sarah.c@sky.net ", "m.scott@dundermifflin.com", "pam.b@gmail..com", "dwight@dunder.com", " invalid_email_here "]
    raw_phones = ["1234567890", "+1 (987) 654-3210", "555-0199", "15550201122", "123 456 7890", "N/A"]
    raw_dates = ["2024-01-15", "02/20/2023", "2022/12/05", "15-08-2021", "2024-05-01", "Invalid Date"]

    data = []
    for i in range(1, 51):
        # Introduce blank rows periodically
        if i in [12, 28, 40]:
            data.append([np.nan, np.nan, np.nan, np.nan, np.nan, np.nan])
            continue

        name = random.choice(raw_names)
        email = random.choice(raw_emails)
        phone = random.choice(raw_phones)
        date = random.choice(raw_dates)
        amount = round(random.uniform(50.0, 1500.0), 2)
        status = random.choice([" COMPLETED ", "pending", "CANCELLED", " Completed "])

        data.append([f" CUST-{100 + (i % 15)} ", name, email, phone, date, amount])

    columns = [" Customer ID ", "Full Name", "Email Address", "Phone Number", "Registration Date", "Account Balance"]
    df = pd.DataFrame(data, columns=columns)

    # Force duplicate rows
    df = pd.concat([df, df.iloc[[2, 5, 10, 15]]], ignore_index=True)

    file_path = os.path.join(OUTPUT_DIR, "01_Messy_Customer_Data.xlsx")
    df.to_excel(file_path, index=False)
    print(f"✅ Generated: {file_path} ({len(df)} rows, includes messy formatting, blanks, and duplicates)")

# ==========================================
# 2. MATCHING & VLOOKUP DATASETS (FILE A & FILE B)
# ==========================================
def generate_matching_datasets():
    # Primary File A (Customers)
    data_a = {
        "CustomerID": [f"CUST-{i:03d}" for i in range(101, 121)],
        "CustomerName": ["Acme Corp", "Globex", "Initech", "Umbrella Corp", "Stark Ind",
                         "Wayne Ent", "Cyberdyne", "Soylent", "Massive Dynamic", "Hooli",
                         "Pied Piper", "Dunder Mifflin", "Aperture", "Black Mesa", "Vehement",
                         "Monsters Inc", "Wonka Ind", "Bluth Co", "Virtucon", "Genco Olive"],
        "Region": ["North", "South", "East", "West"] * 5
    }
    df_a = pd.DataFrame(data_a)
    file_a = os.path.join(OUTPUT_DIR, "02_Dataset_A_Customers.csv")
    df_a.to_csv(file_a, index=False)

    # Secondary File B (Sales Orders - overlapping keys)
    # Includes IDs present in A, plus some missing in A (e.g., CUST-121..125)
    data_b = {
        "OrderID": [f"ORD-{i:04d}" for i in range(5001, 5026)],
        "CustomerID": [f"CUST-{i:03d}" for i in range(105, 126)],
        "OrderValue": [round(random.uniform(100, 5000), 2) for _ in range(25)],
        "OrderStatus": ["Shipped", "Processing", "Delivered", "Cancelled"] * 6 + ["Shipped"]
    }
    df_b = pd.DataFrame(data_b)
    file_b = os.path.join(OUTPUT_DIR, "02_Dataset_B_Orders.xlsx")
    df_b.to_excel(file_b, index=False)

    print(f"✅ Generated: {file_a} (Dataset A for Matching/VLOOKUP)")
    print(f"✅ Generated: {file_b} (Dataset B for Matching/VLOOKUP)")

# ==========================================
# 3. SPLITTING DATASETS (BY COLUMN & BY YEAR)
# ==========================================
def generate_splitting_data():
    departments = ["Sales", "Engineering", "Marketing", "Human Resources", "Finance"]
    start_date = datetime(2021, 1, 1)

    records = []
    for i in range(1, 151):
        emp_id = f"EMP-{i:04d}"
        dept = random.choice(departments)
        random_days = random.randint(0, 365 * 4)  # Spans 2021-2025
        hire_date = start_date + timedelta(days=random_days)
        salary = random.randint(45000, 135000)

        records.append([emp_id, f"Employee_{i}", dept, hire_date.strftime("%Y-%m-%d"), salary])

    df = pd.DataFrame(records, columns=["Employee_ID", "Name", "Department", "Hire_Date", "Salary"])

    file_path = os.path.join(OUTPUT_DIR, "03_MultiYear_Department_Data.xlsx")
    df.to_excel(file_path, index=False)
    print(f"✅ Generated: {file_path} (For testing Split by Column and Split by Year)")

# ==========================================
# 4. BATCH MERGE / CONVERT DATASETS
# ==========================================
def generate_batch_merge_files():
    batch_dir = os.path.join(OUTPUT_DIR, "Batch_CSV_Files_To_Merge")
    os.makedirs(batch_dir, exist_ok=True)

    regions = ["North_Region", "South_Region", "East_Region", "West_Region"]

    for region in regions:
        data = {
            "Region": [region] * 10,
            "Product": [f"Product_{chr(65+i)}" for i in range(10)],
            "Units_Sold": [random.randint(10, 500) for _ in range(10)],
            "Revenue": [round(random.uniform(500, 10000), 2) for _ in range(10)]
        }
        df = pd.DataFrame(data)
        df.to_csv(os.path.join(batch_dir, f"Sales_{region}.csv"), index=False)

    print(f"✅ Generated: 4 CSV files inside folder '{batch_dir}' (For testing File Merging/CSV Conversion)")

if __name__ == "__main__":
    generate_messy_cleaning_data()
    generate_matching_datasets()
    generate_splitting_data()
    generate_batch_merge_files()
    print("\n🎉 All test datasets created successfully in 'sample_test_datasets/'!")