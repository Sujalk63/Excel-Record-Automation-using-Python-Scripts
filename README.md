# Excel Record Automation using Python Scripts

## 📌 Project Overview
This project automates the process of linking employee-specific `.docx` files to their corresponding records in an Excel sheet. By matching **GPN (Global Personal Number)** and **Name** fields, the script dynamically generates hyperlinks that directly point to the correct documents located in a designated directory.

The tool is designed to minimize manual work, enhance data accessibility, and ensure document consistency — ideal for HR teams, operations departments, or any environment dealing with large-scale employee documentation.

---

## 🛠️ What the Script Does
- Reads data from an existing Excel file containing employee records.
- Fetches corresponding `.docx` documents from a specified folder using GPN and Name as identifiers.
- Creates Excel hyperlinks in the relevant row under the “Link” column.
- Saves an updated Excel file with clickable links to each employee's document.

---

## 📦 Libraries Used
- `pandas` – For reading and manipulating the Excel data.
- `os` – For navigating the file system and locating documents.
- `openpyxl` – For writing and formatting Excel files, including hyperlink styling.

---

## ⚙️ Effectiveness
This is a **small yet powerful automation tool** that:
- Saves hours of manual work linking documents row by row.
- Reduces human error in document mapping.
- Can be reused and customized for various Excel-document automation tasks.
- Perfect for office use where document tracking and record-keeping are part of daily operations.

---

## ✅ Ideal Use Cases
- HR departments managing employee files.
- Automating large-scale onboarding/offboarding documentation.
- Centralizing access to project reports, assessments, or verification docs.
- Any workflow that requires quick access to personalized files via Excel sheets.

---

Feel free to customize the paths, naming convention, or extend it with UI/CLI for broader usability.
