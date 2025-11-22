import os
from pathlib import Path
import json

# Import table extraction, storage, and documentation classes
from table_extraction import generate_table_report, extract_formulas
from data_store import process_json_to_postgres
from doc_llm_unique import ExcelFormulaDocGenerator


def main():
    excel_file = r'C:\Users\amsri\PycharmProjects\Excel_parser_cum_data_store\Data\monthly-household-budget - Copy.xlsx'
    outputs_dir = Path('outputs_dir')
    docs_dir = Path('documentation_dir')
    outputs_dir.mkdir(exist_ok=True)
    docs_dir.mkdir(exist_ok=True)

    # 1. Extract tables and formulas, output JSONs to outputs/
    tables_json_path = outputs_dir / f"{Path(excel_file).stem}_tables.json"
    formulas_json_path = outputs_dir / f"{Path(excel_file).stem}_formulas.json"

    generate_table_report(excel_file, tables_json_path)
    extract_formulas(excel_file,table_json_file=tables_json_path,output_json_file=formulas_json_path)

    # 2. Store both JSONs in PostgreSQL
    process_json_to_postgres(
        str(tables_json_path),
        str(formulas_json_path),
        file_name=Path(excel_file).name,
        excel_file=str(excel_file)
    )

    # 3. Generate documentation from formulas JSON
    doc_gen = ExcelFormulaDocGenerator()
    documentation_path = docs_dir / f"{Path(excel_file).stem}_documentation.txt"
    doc_gen.generate_full_documentation(str(formulas_json_path), str(documentation_path))

    print(f"✓ Pipeline complete. Docs in {documentation_path}")


if __name__ == "__main__":
    main()
