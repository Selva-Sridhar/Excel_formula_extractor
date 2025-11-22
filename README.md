## 📘 Excel Formula Extraction & Documentation Pipeline

Automated pipeline for extracting structured table data and formulas from Excel files, generating human-readable documentation using the Gemini API, and storing all metadata in PostgreSQL for auditing and validation.

---

### ✨ Key Features

✔ Extracts **tables** (explicit & implicit) from `.xlsx` and `.xls`
✔ Captures **cell formulas** with references
✔ Generates **documentation reports** using Gemini API
✔ Stores:

* **Table data**
* **Table metadata**
* **Formulas**

in **PostgreSQL** for query & validation
✔ Intermediate output stored as JSON for debugging
✔ Modular architecture for easy scaling

---

### 📂 Project Structure

```
project_root/
├── main.py                    # Pipeline runner
├── table_extraction.py        # Excel table & formula extraction
├── data_store_modified.py     # PostgreSQL storage handlers
├── doc_llm_unique.py          # LLM-based documentation generator
├── outputs/                   # Extracted intermediate JSON files
├── documentation/             # Final generated text reports
├── requirements.txt           # Dependencies
└── .env                       # API Keys and PostgreSQL Config (not versioned)
```

---

### 🚀 Getting Started

#### 1️⃣ Clone the Repository

```bash
git clone https://github.com/Selva-Sridhar/Excel_formula_extractor.git
cd Excel_formula_extractor
```

#### 2️⃣ Create and Activate a Virtual Environment

```bash
python -m venv venv
# Windows:
venv\Scripts\activate
# Linux / macOS:
source venv/bin/activate
```

#### 3️⃣ Set Up `.env` File

Create a `.env` file in the project root and add:

```
PGHOST=localhost             # PostgreSQL server host
PGPORT=5432                  # Default port
PGDATABASE=                  # postgres database name
PGUSER=                      # Username
PGPASSWORD=your_password     # Your password
GOOGLE_API_KEY=your_gemini_key
```

🔐 Never commit `.env` to GitHub!

---

#### 4️⃣ Install Dependencies

```bash
pip install -r requirements.txt
```

---

#### 5️⃣ Run the Pipeline

Edit the input file path in `main.py` and execute:

```bash
python main.py
```

---

### 📌 Outputs

| Output              | Location         | Description                                |
| ------------------- | ---------------- | ------------------------------------------ |
| JSON Extracted Data | `outputs/`       | Sheet-wise structured table dumps          |
| Documentation       | `documentation/` | Human-readable formula explanations        |
| SQL Data            | PostgreSQL       | Stored for validation, reporting, auditing |




