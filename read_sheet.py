import gspread
import pandas as pd

# Step 1: Connect using Service Account (FULL PATH)
client = gspread.service_account(
    filename=r"D:\Coding Works\Visual Studio Code\BIPS IRRELEVANT AUDIT\credentials.json.json"
)

# Step 2: Open Google Sheet by URL
sheet = client.open_by_url(
    "https://docs.google.com/spreadsheets/d/16uZ_W0EVzxIsBUEzVch5UjD01E_Wzhj5D9tOTD61VnY/edit"
)

# Step 3: Select Worksheet (Tab name)
worksheet = sheet.worksheet("Inquiry")

# Step 4: Get all data (handles duplicate column names)
rows = worksheet.get_all_values()
headers = rows[0]
# Make duplicate column names unique: Inquiry Status, Inquiry Status_1, etc.
seen = {}
unique_headers = []
for h in headers:
    if h in seen:
        seen[h] += 1
        unique_headers.append(f"{h}_{seen[h]}")
    else:
        seen[h] = 0
        unique_headers.append(h)

# Step 5: Convert to DataFrame
df = pd.DataFrame(rows[1:], columns=unique_headers)

# Step 6: Show first 5 rows
print("===== FIRST 5 ROWS =====")
print(df.head())

# Step 7: Category-wise count (IMPORTANT 🔥)
if 'Inquiry Status' in df.columns:
    print("\n===== CATEGORY COUNT =====")
    print(df['Inquiry Status'].value_counts())

# Step 8: Save to Excel (MIS Report)
df.to_excel("Inquiry_Report.xlsx", index=False)

print("\nDone! Data exported to Inquiry_Report.xlsx")