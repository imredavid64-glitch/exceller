"""
AI Engine - Gemini API Integration
Handles all AI communication and processing
"""

import os
import requests
import json
import re
from typing import Optional, List, Dict, Tuple


class AIEngine:
    """AI Engine using Google Gemini API"""
    
    def __init__(self, api_key: str = None):
        # API key comes from the GEMINI_API_KEY environment variable or an
        # explicit argument - never hardcode it in source.
        self.api_key = api_key or os.environ.get("GEMINI_API_KEY", "")
        self.api_url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent"
        self.max_tokens = 2048
        
        # System prompt for Excel assistance
        self.system_prompt = """You are an expert Excel assistant named Exceller. You help users with:
- Analyzing spreadsheet data
- Finding errors and inconsistencies
- Suggesting formulas and functions
- Creating charts and visualizations
- Cleaning and transforming data
- Providing statistics and insights
- Writing VBA macros
- Data validation rules

Always provide clear, actionable advice. When showing data, format it nicely.
If a user shares data context, analyze it thoroughly and provide specific recommendations.
Keep responses concise but comprehensive."""
    
    def chat(self, message: str, context: str = "") -> str:
        """Send a chat message and get response"""
        try:
            # Build the prompt with context
            full_prompt = self._build_prompt(message, context)
            
            # Prepare request
            payload = {
                "contents": [
                    {
                        "parts": [
                            {"text": full_prompt}
                        ]
                    }
                ],
                "generationConfig": {
                    "maxOutputTokens": self.max_tokens,
                    "temperature": 0.7,
                    "topP": 0.8,
                    "topK": 40
                },
                "safetySettings": [
                    {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
                    {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
                    {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
                    {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"}
                ]
            }
            
            # Make API request
            response = requests.post(
                f"{self.api_url}?key={self.api_key}",
                headers={"Content-Type": "application/json"},
                json=payload,
                timeout=30
            )
            
            if response.status_code == 200:
                data = response.json()
                return self._extract_response(data)
            else:
                error_msg = f"API Error: {response.status_code}"
                if response.text:
                    try:
                        error_data = response.json()
                        if "error" in error_data:
                            error_msg = error_data["error"].get("message", error_msg)
                    except:
                        pass
                return f"Sorry, I encountered an error: {error_msg}"
                
        except requests.exceptions.Timeout:
            return "Request timed out. Please try again."
        except requests.exceptions.ConnectionError:
            return "Connection error. Please check your internet connection."
        except Exception as e:
            return f"An error occurred: {str(e)}"
    
    def _build_prompt(self, message: str, context: str) -> str:
        """Build the complete prompt with system instructions and context"""
        parts = [
            self.system_prompt,
            "\n\n"
        ]
        
        if context:
            parts.extend([
                "=== CURRENT CONTEXT ===",
                context,
                "\n\n"
            ])
        
        parts.extend([
            "=== USER REQUEST ===",
            message
        ])
        
        return "\n".join(parts)
    
    def _extract_response(self, data: dict) -> str:
        """Extract text response from API response"""
        try:
            if "candidates" in data and len(data["candidates"]) > 0:
                candidate = data["candidates"][0]
                if "content" in candidate and "parts" in candidate["content"]:
                    parts = candidate["content"]["parts"]
                    if parts and "text" in parts[0]:
                        return parts[0]["text"]
            
            return "No response generated. Please try again."
            
        except (KeyError, IndexError) as e:
            return f"Error parsing response: {str(e)}"
    
    def analyze_data(self, data_description: str, data_sample: str) -> str:
        """Analyze data and provide insights"""
        prompt = f"""Analyze this Excel data and provide:
1. Summary of what the data contains
2. Key statistics and patterns
3. Any issues or anomalies found
4. Recommendations for improvement
5. Suggested formulas or analysis

Data Description: {data_description}

Data Sample:
{data_sample}
"""
        return self.chat(prompt)
    
    def suggest_formula(self, description: str, data_context: str = "") -> str:
        """Suggest Excel formula for a task"""
        prompt = f"""Help me write an Excel formula for this task:
{description}

{'Data context: ' + data_context if data_context else ''}

Please provide:
1. The formula
2. Explanation of how it works
3. Example use case
4. Alternative approaches if applicable
"""
        return self.chat(prompt)
    
    def find_errors(self, data: str) -> str:
        """Find errors and inconsistencies in data"""
        prompt = f"""Analyze this data for errors and inconsistencies:
{data}

Look for:
1. Missing values or gaps
2. Duplicate entries
3. Format inconsistencies
4. Type mismatches
5. Logical errors
6. Outliers or anomalies

Provide specific recommendations for each issue found.
"""
        return self.chat(prompt)
    
    def generate_vba(self, task_description: str) -> str:
        """Generate VBA code for a task"""
        prompt = f"""Write VBA code for this Excel task:
{task_description}

Please provide:
1. Complete VBA code
2. Instructions on how to use it
3. Any prerequisites or setup needed
4. Example output if applicable
"""
        return self.chat(prompt)
    
    def suggest_chart(self, data_description: str, data_sample: str) -> str:
        """Suggest appropriate chart type"""
        prompt = f"""Suggest the best chart type for this data:
{data_description}

Data Sample:
{data_sample}

Please provide:
1. Recommended chart type and why
2. How to create it in Excel
3. Customization suggestions
4. Alternative chart options
"""
        return self.chat(prompt)
    
    # =========================================================================
    # FORMULA ERROR FIXER
    # =========================================================================
    
    def fix_formula_error(self, formula: str, error_message: str, 
                          cell_context: str = "", sheet_context: str = "") -> Dict:
        """
        Analyze a formula error and suggest fixes
        
        Returns dict with:
        - error_type: Type of error detected
        - diagnosis: Why the error occurred
        - fixes: List of alternative formulas
        - explanation: Explanation of each fix
        """
        prompt = f"""Analyze this Excel formula error and provide fixes:

ERRORNEOUS FORMULA: {formula}
ERROR MESSAGE: {error_message}
{f'CELL CONTEXT: {cell_context}' if cell_context else ''}
{f'SHEET CONTEXT: {sheet_context}' if sheet_context else ''}

Provide your response in this EXACT JSON format (no markdown, just JSON):
{{
    "error_type": "type of error (e.g., #REF!, #VALUE!, #NAME?, #N/A, #DIV/0!, circular reference, etc.)",
    "diagnosis": "clear explanation of why this error occurred",
    "fixes": [
        {{
            "formula": "corrected formula option 1",
            "explanation": "how this fix works",
            "best_for": "when to use this approach"
        }},
        {{
            "formula": "corrected formula option 2",
            "explanation": "how this fix works",
            "best_for": "when to use this approach"
        }},
        {{
            "formula": "corrected formula option 3",
            "explanation": "how this fix works",
            "best_for": "when to use this approach"
        }}
    ],
    "prevention_tips": ["tip 1", "tip 2"]
}}

IMPORTANT: Provide at least 3 alternative formulas. Consider different approaches like:
- Using different functions (SUMIF vs SUMIFS, VLOOKUP vs INDEX/MATCH vs XLOOKUP)
- Handling errors with IFERROR/IFNA
- Different cell references (absolute vs relative)
- Array formulas vs regular formulas
"""
        response = self.chat(prompt)
        return self._parse_json_response(response)
    
    def suggest_price_calculation(self, current_formula: str, data_columns: List[str],
                                  sample_data: str = "") -> Dict:
        """
        Suggest alternative ways to calculate price/cost
        
        Returns multiple calculation approaches
        """
        columns_str = ", ".join(data_columns) if data_columns else "unknown columns"
        
        prompt = f"""I need alternative ways to calculate price/cost in Excel.

CURRENT FORMULA: {current_formula}
AVAILABLE COLUMNS: {columns_str}
{f'SAMPLE DATA: {sample_data}' if sample_data else ''}

Suggest at least 5 different approaches to calculate the price, including:
1. Basic arithmetic (quantity * unit_price, discounts, etc.)
2. Using SUMPRODUCT for bulk calculations
3. Using VLOOKUP/INDEX-MATCH for price lookups
4. Using PivotTable calculated fields
5. Using Power Query for complex transformations
6. Array formulas for dynamic calculations
7. Using LAMBDA for custom functions (Excel 365)

For each approach provide:
- The exact formula
- When to use it
- Pros and cons
- Example scenario

Format as JSON:
{{
    "approaches": [
        {{
            "name": "approach name",
            "formula": "the formula",
            "description": "how it works",
            "best_for": "when to use",
            "pros": ["pro1", "pro2"],
            "cons": ["con1", "con2"],
            "example": "practical example"
        }}
    ],
    "recommendation": "which approach is best and why"
}}
"""
        response = self.chat(prompt)
        return self._parse_json_response(response)
    
    def auto_fix_formulas(self, formulas: List[Dict]) -> List[Dict]:
        """
        Batch fix multiple formulas at once
        
        Input: List of {"cell": "A1", "formula": "=SUM(B1:B10)", "error": "#REF!"}
        Output: List with fixes applied
        """
        formulas_text = "\n".join([
            f"Cell {f['cell']}: {f['formula']} -> Error: {f.get('error', 'unknown')}"
            for f in formulas
        ])
        
        prompt = f"""Fix these Excel formulas that have errors:

{formulas_text}

For each formula:
1. Identify the error
2. Provide the corrected formula
3. Explain the fix

Format as JSON:
{{
    "fixed_formulas": [
        {{
            "cell": "cell reference",
            "original": "original formula",
            "fixed": "corrected formula",
            "error_type": "type of error",
            "explanation": "what was wrong and how it was fixed"
        }}
    ],
    "summary": "overall summary of issues found"
}}
"""
        response = self.chat(prompt)
        return self._parse_json_response(response)
    
    # =========================================================================
    # VERIFICATION SYSTEM
    # =========================================================================
    
    def verify_formula(self, formula: str, expected_result: str = None,
                       test_data: str = "") -> Dict:
        """
        Verify that a formula is correct and will work
        
        Returns verification result with confidence score
        """
        prompt = f"""Verify this Excel formula is correct:

FORMULA: {formula}
{f'EXPECTED RESULT: {expected_result}' if expected_result else ''}
{f'TEST DATA: {test_data}' if test_data else ''}

Check:
1. Syntax validity - Are all parentheses closed? Correct function names?
2. Reference validity - Do cell references make sense?
3. Logic correctness - Does the formula do what it's supposed to?
4. Edge cases - Will it handle empty cells, errors, etc.?
5. Performance - Is it efficient for large datasets?

Format as JSON:
{{
    "is_valid": true/false,
    "confidence_score": 0-100,
    "syntax_check": {{
        "valid": true/false,
        "issues": ["issue1", "issue2"]
    }},
    "reference_check": {{
        "valid": true/false,
        "issues": ["issue1"]
    }},
    "logic_check": {{
        "valid": true/false,
        "explanation": "does it do what intended"
    }},
    "edge_cases": {{
        "handles_empty": true/false,
        "handles_errors": true/false,
        "handles_large_data": true/false
    }},
    "improvements": ["suggestion1", "suggestion2"],
    "verified_formula": "improved version if needed",
    "warnings": ["warning1", "warning2"]
}}
"""
        response = self.chat(prompt)
        return self._parse_json_response(response)
    
    def test_formula_with_data(self, formula: str, test_cases: List[Dict]) -> Dict:
        """
        Test a formula with multiple test cases
        
        Input: List of {"input": {...}, "expected": result}
        Output: Test results with pass/fail for each case
        """
        test_cases_text = "\n".join([
            f"Test {i+1}: Input={tc.get('input', {})}, Expected={tc.get('expected', 'N/A')}"
            for i, tc in enumerate(test_cases)
        ])
        
        prompt = f"""Test this Excel formula with the given test cases:

FORMULA: {formula}

TEST CASES:
{test_cases_text}

For each test case:
1. Calculate what the formula would return
2. Compare with expected result
3. Note any issues

Format as JSON:
{{
    "formula": "the formula being tested",
    "test_results": [
        {{
            "test_number": 1,
            "input": {{...}},
            "calculated_result": "what formula returns",
            "expected_result": "what was expected",
            "passed": true/false,
            "notes": "any observations"
        }}
    ],
    "summary": {{
        "total_tests": 5,
        "passed": 4,
        "failed": 1,
        "pass_rate": "80%"
    }},
    "verdict": "PASS/FAIL/PARTIAL",
    "issues_found": ["issue1"]
}}
"""
        response = self.chat(prompt)
        return self._parse_json_response(response)
    
    def verify_vba_code(self, code: str, description: str = "") -> Dict:
        """
        Verify VBA code is correct and will work
        
        Returns verification with syntax check and suggestions
        """
        prompt = f"""Verify this VBA code is correct:

VBA CODE:
{code}

{f'DESCRIPTION: {description}' if description else ''}

Check:
1. Syntax validity
2. Variable declarations
3. Object references
4. Error handling
5. Best practices

Format as JSON:
{{
    "is_valid": true/false,
    "confidence_score": 0-100,
    "syntax_errors": [
        {{"line": 5, "error": "error description", "fix": "how to fix"}}
    ],
    "warnings": [
        {{"line": 10, "warning": "warning description", "suggestion": "improvement"}}
    ],
    "best_practices": ["tip1", "tip2"],
    "verified_code": "corrected version if needed",
    "explanation": "overall assessment"
}}
"""
        response = self.chat(prompt)
        return self._parse_json_response(response)
    
    def validate_output(self, formula_or_code: str, output_description: str,
                        actual_output: str = None) -> Dict:
        """
        Validate that the AI's output matches expectations
        
        This is the self-check mechanism - verifies its own recommendations
        """
        prompt = f"""Validate this Excel solution:

SOLUTION: {formula_or_code}
EXPECTED BEHAVIOR: {output_description}
{f'ACTUAL OUTPUT (if tested): {actual_output}' if actual_output else ''}

Perform a self-check:
1. Does the solution actually solve the problem?
2. Are there any logical errors?
3. Will it work in all scenarios?
4. Is it the most efficient approach?
5. Are there any edge cases not handled?

Format as JSON:
{{
    "is_correct": true/false,
    "confidence_score": 0-100,
    "checks": {{
        "solves_problem": true/false,
        "no_logical_errors": true/false,
        "works_all_scenarios": true/false,
        "efficient": true/false,
        "handles_edge_cases": true/false
    }},
    "issues": [
        {{"type": "critical/warning/info", "description": "issue description"}}
    ],
    "alternative_solution": "better solution if exists",
    "final_verdict": "APPROVED/NEEDS_REVISION/REJECTED",
    "revision_notes": "what to change if needed"
}}
"""
        response = self.chat(prompt)
        return self._parse_json_response(response)
    
    def _parse_json_response(self, response: str) -> Dict:
        """Parse JSON from AI response, handling markdown code blocks"""
        try:
            # Remove markdown code blocks if present
            cleaned = re.sub(r'```json?\s*', '', response)
            cleaned = re.sub(r'```\s*$', '', cleaned)
            cleaned = cleaned.strip()
            
            # Try to find JSON object in the response
            json_match = re.search(r'\{[\s\S]*\}', cleaned)
            if json_match:
                return json.loads(json_match.group())
            
            # If no JSON found, return raw response
            return {"raw_response": response, "parse_error": "Could not extract JSON"}
            
        except json.JSONDecodeError as e:
            return {"raw_response": response, "parse_error": str(e)}
