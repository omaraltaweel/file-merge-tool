
# 📎 Standard Materials Import Tool (Streamlit App)

This tool provides a secure and user-friendly interface to merge multiple Excel files containing structured data into one clean, validated, and formatted Excel sheet.

---

## ✅ Features

### 🔐 Password Protection
- Access is restricted with a password prompt at the start of the session.
- Prevents unauthorized use of the tool.

### 📤 Excel File Upload
- Drag and drop interface for uploading `.xlsx`, `.xlsm`, and optionally `.xls` files.
- Supports multiple files in one session.
- Upload buffer managed with session state.

### 🧠 Case-Insensitive Header and Sheet Detection
- Automatically detects the `"Standard Materials"` sheet, even if it’s named `standard materials`, `STANDARD MATERIALS`, etc.
- Column headers are validated case-insensitively against a predefined template.

### 🔍 Validation Feedback
- If a file is missing the required sheet or has incorrect headers, it's reported with detailed error messages.

### 🧾 Data Cleaning
- Trims all cell text (removes leading/trailing spaces).
- Removes fully empty columns from unwanted predefined fields.

### 🟨 Highlighting
- Duplicate `Material_ID` values are highlighted in **red**.
- Columns that should be empty but contain data (based on config) are highlighted in **light red**.

### 💬 Source File Tracking
- Adds the original file name as a **comment** on each `Material_ID` cell.

### 🔽 Output Sorting
- Output Excel data is automatically sorted by `Material_ID` in ascending order.

### 📥 Download Output
- Single-click download of the merged, formatted Excel file.

---

## 📦 Dependencies

Make sure your environment includes:

```
pandas
openpyxl
streamlit >= 1.25
```

If using `.xls` files:
```
xlrd
```

## ✨ Example Use Case

- Clean and merge material specification data from multiple suppliers
- Consolidate engineering BOM data from various templates
- Validate standardised sheets before reporting or importing to ERP

---

Created with ❤️ using Streamlit and Python
