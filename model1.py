import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# File paths
excel_file = r"C:\Users\Sujal Karmakar\Desktop\Desktop\automate\automate.xlsx"
docs_folder = r"C:\Users\Sujal Karmakar\Desktop\Desktop\automate\docs"
output_file = r"C:\Users\Sujal Karmakar\Desktop\Desktop\automate\updated_automate.xlsx"

# Reading the Excel file
df = pd.read_excel(excel_file)

# Ensure the 'Link' column exists
df["Link"] = ""

for index, row in df.iterrows():
    gpn = str(row["GPN"])
    name = row["Name"].strip().lower()  # Match file naming convention
    
    expected_filename = f"{gpn}_{name}.docx"
    file_path = os.path.join(docs_folder, expected_filename)
    
    if os.path.exists(file_path):
        df.at[index, "Link"] = f'=HYPERLINK("{file_path}", "link")'

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df.to_excel(writer, index=False)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    
    # Auto-format the 'Link' column to be clickable
    for row in range(2, len(df) + 2): 
        worksheet[f"C{row}"].style = "Hyperlink"

print("Created and Updated Excel file saved successfully!")
