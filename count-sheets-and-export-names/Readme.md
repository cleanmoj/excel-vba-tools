# Excel Sheet Name Counter & Extractor

This VBA tool adds a feature to Excel that allows users to **count & extract sheet names** from a workbook and save them in various formats, which is **not available by default** in Excel.

## 🔍 Why This Tool?

Excel does not provide a built-in way to extract sheet names from a workbook or count the number of sheets. This can be cumbersome for large workbooks with many sheets. This tool solves that problem by offering an easy way to extract and save sheet names in multiple formats.

## 🚀 How to Use

1. **Open your Excel file** (macro-enabled: `.xlsm`).
2. **Press `ALT + F11`** to open the VBA Editor.
3. **Import the `.bas` file** from this tool into your project.
4. Close the VBA Editor.
5. **Press `ALT + F8`** to run the macro named `CountSheetsAndExportTheirNames`.

## 📁 Files Included

- `CountSheetsAndExportNames.bas` — The main VBA module to extract sheet names.


## 🎬 Demonstration

![Demo](assets/count-sheet-demo.gif)

## 📌 Notes

- The tool allows you to save sheet names and the total count to the following formats:
  - **Excel (.xlsx)**
  - **CSV (.csv)**
  - **Text (.txt)**
- The total count of sheets is included in all outputs.
- You can install this macro in your personal macro workbook (`Personal.xlsb`) to make it available in all Excel files.
- Requires macros to be enabled.

---

Feel free to customize or improve the code to better suit your needs. Contributions are welcome!
