# SaveAsCSVWithoutOpeningIt

---

## 🧾 AutoSaveSheetAsUTF8CSVWithTimestamp

### 🔍 Description
This macro exports the **active worksheet** as a UTF-8 encoded `.csv` file.  
It automatically generates a timestamped file name and saves it in the same directory as the source workbook.

---

### 🛠️ Features
- Automatically names the file like: `WorkbookName_Export_20240625_1730.csv`
- No prompts — fully automatic
- Preserves original workbook
- Prevents filename errors by using timestamp-safe formats

---

### 📦 Output Example

If your workbook is named `MonthlyReport.xlsm` and today is June 25, 2025 at 5:30 PM, the saved file will be:

```text
MonthlyReport_Export_20250625_1730.csv
```

---

### ✅ How to Use

1. Import `AutoSaveSheetAsUTF8CSVWithTimestamp.bas` into your VBA project.
2. Run the macro from `ALT + F8` or assign it to a button.
3. The active sheet will be saved as a UTF-8 `.csv` file in the same folder.

---

### 📋 Requirements

- Excel 2016 or later
- Macros enabled (`.xlsm` file recommended)

---

### 💡 Notes

- Only the **active sheet** is exported.
- The CSV file is not opened after creation.
- If you need to export multiple sheets, let me know and I can help extend the functionality.

---

## 📄 License

MIT License — free to use, modify, and redistribute.

---

## 🙌 Credits

Made by Mohamed El-ansary — inspired by the need to version Excel workbooks without version control software.

---
