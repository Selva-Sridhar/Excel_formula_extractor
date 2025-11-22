import json
import google.generativeai as genai
from typing import List, Dict, Any
import os
from collections import defaultdict
from pathlib import Path

# Configure Gemini API
genai.configure(api_key=os.environ.get('GOOGLE_API_KEY'))


class ExcelFormulaDocGenerator:
    def __init__(self, api_key: str = None):
        """Initialize the documentation generator with Gemini API."""
        if api_key:
            genai.configure(api_key=api_key)
        self.model = genai.GenerativeModel('gemini-2.5-flash')

    def load_formulas(self, json_file_path: str) -> List[Dict[str, Any]]:
        """Load formulas from JSON file."""
        with open(json_file_path, 'r', encoding='utf-8') as f:
            return json.load(f)

    def group_formulas_by_pattern(self, formulas: List[Dict[str, Any]]) -> Dict[str, List[Dict[str, Any]]]:
        """
        Group formulas by their pattern to identify redundant formulas.
        Returns dict: {readable_formula: [list of formula objects with that pattern]}
        """
        pattern_groups = defaultdict(list)

        for formula_obj in formulas:
            formula = formula_obj.get('formula', '')
            readable = formula_obj.get('readable_formula', '')

            # Use readable formula as primary pattern, fall back to formula
            if readable and readable != formula:
                pattern_key = readable
            else:
                pattern_key = formula

            pattern_groups[pattern_key].append(formula_obj)

        return dict(pattern_groups)

    def create_unique_formula_summary(self, pattern_groups: Dict[str, List[Dict[str, Any]]]) -> List[Dict[str, Any]]:
        """
        Create a summary of unique formulas with their occurrences.
        """
        unique_formulas = []

        for pattern, instances in pattern_groups.items():
            # Get representative instance
            representative = instances[0]

            # Collect all cells where this pattern appears
            cells = [inst['cell'] for inst in instances]

            unique_formulas.append({
                'pattern': pattern,
                'formula_example': representative['formula'],
                'readable_formula': representative.get('readable_formula', pattern),
                'cells': cells,
                'occurrence_count': len(instances),
                'dependencies': representative.get('dependencies', []),
                'context': representative.get('context', {})
            })

        return unique_formulas

    def group_by_sheet(self, formulas: List[Dict[str, Any]]) -> Dict[str, List[Dict[str, Any]]]:
        """Group formulas by sheet name."""
        sheets = {}
        for formula in formulas:
            sheet_name = formula.get('context', {}).get('sheet', 'Unknown')
            if sheet_name not in sheets:
                sheets[sheet_name] = []
            sheets[sheet_name].append(formula)
        return sheets

    def create_prompt(self, unique_formulas: List[Dict[str, Any]], sheet_name: str) -> str:
        """Create a detailed prompt for Gemini to document formulas."""

        prompt = f"""You are an Excel formula documentation expert. Analyze the following UNIQUE Excel formulas from the sheet "{sheet_name}".

        CRITICAL INSTRUCTIONS:
        1. Use the 'readable_formula' field to understand what the formula does in plain English
        2. When describing dependencies, use meaningful column names from 'readable_formula' (like "Actual", "Budget") NOT cell references (like C5, B5)
        3. For mathematical notation, use the readable column names: "Actual - Budget = Difference" NOT "C5 - B5 = D5"
        4. Generate output as PLAIN TEXT with clear sections and headings using text formatting (not markdown)
        5. Use simple text formatting: CAPITALIZED HEADINGS, numbered lists, and clear spacing
        6. Do NOT use any markdown syntax (no #, **, -, etc.)
        7. Group formulas by their PATTERN (e.g., all subtraction formulas together, all division formulas together)

        Excel Function Reference: https://support.microsoft.com/en-us/office/excel-functions-alphabetical-b3944572-255d-4efb-bb96-c6d90033e188

        UNIQUE FORMULAS DATA:
        {json.dumps(unique_formulas, indent=2)}

        Create comprehensive documentation with EXACTLY these 4 parts:

        ================================================================================
        PART 1: FULL OVERVIEW
        ================================================================================

        1.1 Sheet Purpose
           - What is the primary purpose of this sheet?
           - What business problem does it solve?

        1.2 Sheet Structure
           - How is the data organized?
           - What are the main sections/tables?

        1.3 Key Calculations
           - What are the most important calculations performed?
           - List 3-5 critical formulas and their purpose

        1.4 Data Flow
           - How does data flow through the sheet?
           - Input sources -> Processing -> Outputs
           - Which cells feed into which calculations?

        ================================================================================
        PART 2: FORMULA DOCUMENTATION (GROUPED BY PATTERN)
        ================================================================================

        Group formulas by their mathematical pattern:
        - Pattern 1: SUBTRACTION (=[COL1]-[COL2])
        - Pattern 2: ADDITION/SUM
        - Pattern 3: DIVISION
        - Pattern 4: MULTIPLICATION
        - Pattern 5: TABLE REFERENCES
        - Pattern 6: SUBTOTAL/AGGREGATE FUNCTIONS
        - Pattern 7: TEXT/CONCATENATION
        - Pattern 8: COMPLEX/NESTED FORMULAS

        For each pattern group:

          Pattern: [Pattern Name - e.g., "Subtraction: Value1 - Value2"]

          Formula Examples:
            - Show the readable formula: "Actual - Budget = Difference"
            - Mathematical notation with meaningful names (NOT cell references)

          Occurrence Count: [X formulas follow this pattern]

          Applied to Cells: [List all cells where this pattern appears]

          Purpose:
            - What does this pattern calculate?
            - Why is it used multiple times?

          Dependencies:
            - What columns/tables does it depend on? (use readable names)
            - Relationship explanation in plain language

          Business Context:
            - What does this mean for the user/business?
            - Any special considerations or edge cases?

          ---

        ================================================================================
        PART 3: DEPENDENCY ANALYSIS
        ================================================================================

        3.1 Formula Dependencies
           - Which formulas depend on other formulas?
           - Create a dependency map showing relationships
           - Example: "Cell I5 depends on G5 and H5, which in turn depend on Table2"

        3.2 Table Dependencies
           - Which formulas reference structured tables?
           - What tables are most heavily used?

        3.3 Cross-Reference Map
           - Show the flow: Source Data -> Intermediate Calculations -> Final Results
           - Identify calculation chains

        3.4 Circular References & Issues
           - Any potential circular references?
           - Any broken or suspicious dependencies?

        ================================================================================
        PART 4: INSIGHTS & ANALYSIS
        ================================================================================

        4.1 Formula Patterns Identified
           - What are the dominant calculation patterns?
           - Are there repeated calculation logic?
           - Pattern frequency analysis

        4.2 Best Practices Observed
           - What is done well in this sheet?
           - Good use of structured references, named ranges, etc.

        4.3 Potential Issues
           - Any hardcoded values that should be parameters?
           - Overly complex formulas that could be simplified?
           - Maintainability concerns?

        4.4 Key Findings
           - Top 5 most important insights about this sheet
           - Summary of calculation methodology
           - Overall assessment of formula quality and organization

        4.5 Recommendations
           - Suggestions for improvement
           - Formula optimization opportunities
           - Documentation or naming improvements

        ================================================================================

        Format as clear, readable plain text with these exact 4 parts. Use section dividers as shown.
        """
        return prompt

    def generate_sheet_documentation(self, sheet_formulas: List[Dict[str, Any]], sheet_name: str) -> str:
        """Generate documentation for a single sheet with unique formulas."""

        # Group by pattern to find unique formulas
        pattern_groups = self.group_formulas_by_pattern(sheet_formulas)
        unique_formulas = self.create_unique_formula_summary(pattern_groups)

        print(f"  Found {len(sheet_formulas)} total formulas, {len(unique_formulas)} unique patterns")

        prompt = self.create_prompt(unique_formulas, sheet_name)

        try:
            response = self.model.generate_content(prompt)
            return response.text
        except Exception as e:
            return f"Error generating documentation for {sheet_name}: {str(e)}"

    def generate_full_documentation(self, json_file_path: str, output_file: str = "formula_documentation.txt"):
        """Generate complete documentation for all sheets in the JSON file."""

        print("Loading formulas...")
        formulas = self.load_formulas(json_file_path)

        print("Grouping formulas by sheet...")
        sheets = self.group_by_sheet(formulas)

        # Create the main documentation
        documentation = "=" * 80 + "\n"
        documentation += "EXCEL FORMULA DOCUMENTATION\n"
        documentation += "=" * 80 + "\n\n"
        documentation += f"Generated from: {json_file_path}\n"
        documentation += f"Total Sheets: {len(sheets)}\n"
        documentation += f"Total Formulas: {len(formulas)}\n\n"
        documentation += "=" * 80 + "\n\n"

        # Add table of contents
        documentation += "TABLE OF CONTENTS\n"
        documentation += "-" * 80 + "\n\n"
        for i, sheet_name in enumerate(sheets.keys(), 1):
            documentation += f"{i}. {sheet_name}\n"
        documentation += "\n" + "=" * 80 + "\n\n"

        # Generate documentation for each sheet
        for sheet_name, sheet_formulas in sheets.items():
            print(f"\nProcessing sheet: {sheet_name}...")
            sheet_doc = self.generate_sheet_documentation(sheet_formulas, sheet_name)

            documentation += "\n" + "=" * 80 + "\n"
            documentation += f"SHEET: {sheet_name}\n"
            documentation += "=" * 80 + "\n\n"
            documentation += sheet_doc
            documentation += "\n\n" + "=" * 80 + "\n\n"

        # Write to file
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(documentation)

        print(f"\n✓ Documentation generated successfully: {output_file}")
        return documentation
