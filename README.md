# Block Data Digitization Pipeline

Automated Python data processing toolkit designed to clean, transform, and structure raw multi-district block operational data into standardized Excel workbooks.

## 🛠️ Key Capabilities
- **Automated Ingestion:** Standardizes disparate district spreadsheets across administrative units.
- **Hygiene & Validation:** Cleans missing identifiers, trims whitespace, standardizes naming conventions, and deduplicates records.
- **Workbook Generation:** Employs `pandas` and `openpyxl` to generate formatted `.xlsx` workbooks with summary sheets.

## 🚀 Getting Started
```bash
# Clone the repository
git clone https://github.com/nadeemalam178/Block-Data-Digitization.git
cd Block-Data-Digitization

# Install dependencies
pip install pandas openpyxl

# Execute generation pipeline
python generate_all_workbooks.py
```

## ⚖️ License
MIT License.
