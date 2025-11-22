# Excel Formula Extraction and Documentation Pipeline

Automated system for extracting tables and formulas from Excel files, generating human-readable documentation, and storing all relevant metadata in PostgreSQL.

---

## Features

- Extracts explicit and implicit tables, along with all formulas, from `.xlsx` and `.xls` files
- Generates documentation summarizing formulas and calculations using Gemini API
- Stores table data, table metadata, and formulas in PostgreSQL for validation and team audit
- Saves intermediate JSON files in `outputs/` for review and debugging

---

## Directory Structure

project_root/
├── main.py
├── table_extraction.py
├── data_store_modified.py
├── doc_llm_unique.py
├── outputs/ # Intermediate JSON files
├── documentation/ # Final documentation
├── requirements.txt
└── .env # Secrets/config (not added to Git)


---

## Getting Started

### 1. Clone the Repository
git clone (https://github.com/Selva-Sridhar/Excel_formula_extractor.git)
cd Excel_formula_extractor

### 2. Create and Activate a Virtual Environment
python -m venv venv
source venv/bin/activate # Windows: venv\Scripts\activate


### 3. Prepare the `.env` File

Create a `.env` file in your project root:
PGHOST= # IP or address of your PostgreSQL server
PGPORT= # Usually 5432
PGDATABASE= # Database name
PGUSER= # PostgreSQL user
PGPASSWORD= # Password for the user
GOOGLE_API_KEY=# Gemini API key


### 4. Install Python Dependencies

pip install -r requirements.txt


### 5. Run the Pipeline

Edit `main.py` to set the name of your Excel file, then run:

python main.py


- Documentation will be generated as `{excel_file}_documentation.txt` in `documentation/`
- Extracted tables and formulas JSON files will appear in `outputs/` (for validation)
- Data is stored in your configured PostgreSQL database


