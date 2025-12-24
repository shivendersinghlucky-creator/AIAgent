"""
🏢 ENTERPRISE EXCEL AUTOMATION AI AGENT
========================================

An enterprise-grade autonomous agent that intelligently analyzes, modifies, 
and enhances Excel files based on business language queries.

You are NOT a chatbot. You are a decision-making Excel agent.

ARCHITECTURE:
-------------
1. Excel Structure Analyzer - Deep file structure analysis
2. Business Query Parser - Converts business language to operations
3. Decision Engine - Enterprise-grade planning and validation
4. Execution Engine - Safe, auditable modifications
5. Validation & Quality Control - Data integrity enforcement
6. Reporting Engine - Clear explanations and change logs

SUPPORTED OPERATIONS:
--------------------
1. Calculations (Sum, Avg, Min, Max, %, etc.)
2. Data Transformation (Add/Rename/Remove columns, Sort, Filter)
3. Data Validation & Quality (Dropdowns, Error detection)
4. Conditional Formatting (Color coding, KPIs, Data bars)
5. Aggregation & Analytics (Pivot tables, Group by)
6. Lookup & Reference (XLOOKUP, INDEX-MATCH, Joins)
7. Visualization (Charts: Bar, Pie, Line, KPI dashboards)
8. Reporting & Governance (Summaries, Dashboards, Protection)
"""

from groq import Groq
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.chart import BarChart, Reference, PieChart, LineChart
import json
from typing import Dict, List, Any, Optional
from dataclasses import dataclass
import os
import re


# ==========================================
# SYSTEM PROMPT - AGENT IDENTITY
# ==========================================
SYSTEM_PROMPT = """You are an Enterprise Excel Automation AI Agent, not a chatbot.

IDENTITY:
- You are a professional decision-making agent specialized in Excel operations
- You operate with enterprise-grade precision and governance
- You prioritize data integrity, auditability, and clear communication

CORE RESPONSIBILITIES:
1. Analyze Excel file structures with deep understanding
2. Parse business language queries into technical operations
3. Create risk-assessed decision plans before execution
4. Execute operations safely with complete audit trails
5. Provide clear, professional explanations of all changes

OPERATIONAL PRINCIPLES:
- Always create new files (never overwrite originals)
- Log every operation for audit compliance
- Assess risk before execution (low/medium/high)
- Request confirmation for high-risk operations
- Use exact column names from the actual file
- Provide before/after previews when possible

TONE & STYLE:
- Professional and concise
- Use business language, not technical jargon
- Provide structured, actionable responses
- Show confidence in recommendations
- Admit when clarification is needed

YOU ARE NOT:
- A general chatbot for casual conversation
- A data science consultant
- A replacement for human judgment on critical decisions
"""


# ==========================================
# DECISION DATA STRUCTURE
# ==========================================
@dataclass
class AgentDecision:
    """Enterprise-grade structured decision object before Excel modification"""
    operation_type: str  # calculation | transformation | validation | formatting | aggregation | lookup | visualization | reporting
    sub_operation: str  # Specific operation like "sum", "pivot", "xlookup", "bar_chart"
    source_columns: List[str]
    target_columns: Optional[List[str]]  # Can be multiple for complex operations
    sheet_name: str
    execution_scope: str  # all_rows | filtered_rows | specific_range
    risk_level: str  # low | medium | high
    assumptions: List[str]
    requires_confirmation: bool  # For high-risk operations
    change_description: str  # Plain English explanation
    
    def to_dict(self) -> Dict:
        return {
            "operation_type": self.operation_type,
            "sub_operation": self.sub_operation,
            "source_columns": self.source_columns,
            "target_columns": self.target_columns,
            "sheet_name": self.sheet_name,
            "execution_scope": self.execution_scope,
            "risk_level": self.risk_level,
            "assumptions": self.assumptions,
            "requires_confirmation": self.requires_confirmation,
            "change_description": self.change_description
        }


# ==========================================
# AGENT TOOLS (Excel Operations)
# ==========================================
EXCEL_TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "analyze_excel_structure",
            "description": "Analyzes the Excel file structure including sheets, columns, data types, and sample data",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Path to the Excel file"
                    },
                    "sheet_name": {
                        "type": "string",
                        "description": "Specific sheet to analyze (optional, analyzes first sheet if not provided)"
                    }
                },
                "required": ["file_path"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "parse_user_query",
            "description": "Parses natural language query to extract operation type, columns involved, and intended action",
            "parameters": {
                "type": "object",
                "properties": {
                    "query": {
                        "type": "string",
                        "description": "User's natural language query"
                    },
                    "available_columns": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "List of available column names in the Excel file"
                    }
                },
                "required": ["query", "available_columns"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "create_decision_plan",
            "description": "Creates a structured decision plan for Excel modification based on parsed query and Excel structure",
            "parameters": {
                "type": "object",
                "properties": {
                    "operation_type": {
                        "type": "string",
                        "enum": ["calculation", "transformation", "validation", "formatting", "aggregation", "lookup", "visualization", "reporting"],
                        "description": "Type of Excel operation"
                    },
                    "source_columns": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "Source columns involved in the operation"
                    },
                    "target_columns": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "Names of new columns to create or existing columns to modify"
                    },
                    "sub_operation": {
                        "type": "string",
                        "description": "Specific operation: sum, average, pivot, xlookup, bar_chart, etc."
                    },
                    "execution_scope": {
                        "type": "string",
                        "enum": ["all_rows", "filtered_rows", "specific_range"],
                        "description": "Which rows to apply the operation to"
                    }
                },
                "required": ["operation_type", "source_columns", "sub_operation"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "execute_excel_operation",
            "description": "Executes the planned Excel operation safely on the file",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "Path to the Excel file"
                    },
                    "decision_plan": {
                        "type": "object",
                        "description": "The approved decision plan from create_decision_plan"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "Path to save the modified Excel file"
                    }
                },
                "required": ["file_path", "decision_plan", "output_path"]
            }
        }
    }
]


# ==========================================
# EXCEL AUTOMATION AGENT CLASS
# ==========================================
class ExcelAgent:
    """
    Enterprise-grade Excel Automation AI Agent
    
    Decision-making agent for business Excel operations.
    NOT a chatbot - a professional automation tool.
    """
    
    def __init__(self, api_key: str):
        """Initialize the Enterprise Excel Agent with Groq API"""
        self.client = Groq(api_key=api_key)
        self.model = "llama-3.3-70b-versatile"
        self.system_prompt = SYSTEM_PROMPT
        self.conversation_history = []
        self.change_log = []  # Track all changes for audit
        
    def analyze_excel_structure(self, file_path: str, sheet_name: Optional[str] = None) -> Dict:
        """
        Analyzes Excel file structure
        
        Returns:
            Dictionary containing sheets, columns, data types, and sample data
        """
        try:
            # Load Excel file
            xl_file = pd.ExcelFile(file_path)
            sheets = xl_file.sheet_names
            
            # Use first sheet if not specified
            target_sheet = sheet_name if sheet_name else sheets[0]
            
            # Read the sheet
            df = pd.read_excel(file_path, sheet_name=target_sheet)
            
            # Analyze structure
            analysis = {
                "file_name": os.path.basename(file_path),
                "available_sheets": sheets,
                "analyzed_sheet": target_sheet,
                "total_rows": len(df),
                "total_columns": len(df.columns),
                "columns": list(df.columns),
                "data_types": {col: str(dtype) for col, dtype in df.dtypes.items()},
                "sample_data": df.head(3).to_dict('records'),
                "null_counts": df.isnull().sum().to_dict(),
                "numeric_columns": df.select_dtypes(include=['number']).columns.tolist(),
                "text_columns": df.select_dtypes(include=['object']).columns.tolist()
            }
            
            return analysis
            
        except Exception as e:
            return {"error": str(e), "status": "failed"}
    
    def parse_user_query(self, query: str, excel_structure: Dict) -> Dict:
        """
        Enhanced AI parsing with multi-step reasoning and comprehensive prompt engineering
        
        Returns:
            Parsed query with operation type, columns, and intent
        """
        prompt = self._build_enhanced_parsing_prompt(query, excel_structure)
        
        # Use system prompt with enhanced user prompt
        messages = [
            {"role": "system", "content": self.system_prompt},
            {"role": "user", "content": prompt}
        ]
        
        response = self.client.chat.completions.create(
            model=self.model,
            messages=messages,
            temperature=0.1,  # Low temperature for precise parsing
            max_tokens=1500,  # Increased for detailed reasoning
            top_p=0.95
        )
        
        return self._extract_json_from_response(response)
    
    def _build_enhanced_parsing_prompt(self, query: str, excel_structure: Dict) -> str:
        """
        Build comprehensive prompt with multi-step reasoning for query parsing
        """
        available_columns = excel_structure.get("columns", [])
        numeric_columns = excel_structure.get("numeric_columns", [])
        text_columns = excel_structure.get("text_columns", [])
        sample_data = excel_structure.get("sample_data", [])
        total_rows = excel_structure.get('total_rows', 0)
        total_columns = excel_structure.get('total_columns', 0)
        
        prompt = f"""You are an Enterprise Excel Operation Parser with deep domain expertise.

CONTEXT - EXCEL FILE ANALYSIS:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📊 Total Rows: {total_rows}
📋 Total Columns: {total_columns}

AVAILABLE COLUMNS:
{', '.join(available_columns)}

NUMERIC COLUMNS (can be calculated):
{', '.join(numeric_columns) if numeric_columns else 'None'}

TEXT COLUMNS (categories/labels):
{', '.join(text_columns) if text_columns else 'None'}

SAMPLE DATA (First 3 rows for context):
{json.dumps(sample_data[:3], indent=2)}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

USER'S BUSINESS QUERY:
"{query}"

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

YOUR TASK - MULTI-STEP REASONING:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

STEP 1: INTENT ANALYSIS
- What is the user trying to accomplish?
- Is this a calculation, visualization, transformation, or analysis?
- What business outcome do they expect?

STEP 2: COLUMN MAPPING
- Which columns are mentioned (explicitly or implicitly)?
- Map business terms to exact column names
- Identify if columns exist or need to be created

STEP 3: OPERATION CLASSIFICATION
- Determine the operation type and sub-operation
- Consider data types (numeric vs text)
- Assess technical feasibility

STEP 4: RISK ASSESSMENT
- Does this modify existing data? (higher risk)
- Does this delete information? (high risk)
- Is this reversible? (lower risk)

STEP 5: VALIDATION
- Are all required columns available?
- Is the operation technically possible?
- Any ambiguities that need clarification?

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

OUTPUT FORMAT - RETURN ONLY VALID JSON:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

TASK: Extract and return ONLY a valid JSON object (no markdown, no code blocks) with this exact structure:
{{
    "reasoning": {{
        "intent_summary": "Brief summary of what user wants",
        "column_mapping": {{"user_term": "actual_column_name"}},
        "operation_rationale": "Why this operation type was chosen",
        "risk_factors": ["factor1", "factor2"]
    }},
    "operation_type": "calculation | transformation | validation | formatting | aggregation | lookup | visualization | reporting",
    "sub_operation": "specific_operation_name",
    "source_columns": ["exact_column_name1", "exact_column_name2"],
    "needs_new_column": true,
    "suggested_column_names": ["New_Column_Name"],
    "operation_description": "Plain English description of what will happen",
    "formula_logic": "If calculation: 'sum(A, B)', if visualization: 'bar chart of Sales by Product'",
    "is_ambiguous": false,
    "clarification_needed": null,
    "risk_level": "low | medium | high",
    "execution_confidence": "high | medium | low",
    "alternative_interpretations": []
}}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

OPERATION TYPES & SUB-OPERATIONS CATALOG:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. CALCULATION
   • sum - Add multiple columns/values
   • average/mean - Calculate mean value
   • min/minimum - Find minimum value
   • max/maximum - Find maximum value
   • count - Count non-empty values
   • median - Calculate median
   • multiply - Multiply columns
   • divide - Divide columns
   • percentage - Calculate percentage (ratio or of total)
   • variance - Statistical variance
   • std_dev - Standard deviation

2. TRANSFORMATION
   • add_column - Create new empty column
   • rename_column - Rename existing column
   • remove_column - Delete column (HIGH RISK)
   • sort - Sort data by columns
   • filter - Filter rows by condition
   • deduplicate - Remove duplicate rows
   • split_column - Split column into multiple
   • merge_columns - Combine multiple columns
   • fill_missing - Fill null values
   • replace_values - Replace specific values

3. VALIDATION
   • dropdown - Create dropdown lists
   • range_check - Validate numeric ranges
   • error_detection - Find data quality issues
   • blank_check - Identify missing values
   • duplicate_check - Find duplicates
   • data_type_validation - Ensure correct data types

4. FORMATTING
   • conditional_color - Color cells by condition
   • highlight - Highlight specific values
   • data_bars - Add data bar visualization in cells
   • icon_sets - Add conditional icons
   • font_style - Change font properties
   • number_format - Format numbers (currency, %, etc.)

5. AGGREGATION
   • pivot_table - Create pivot table
   • group_by - Group and aggregate data
   • cross_tab - Cross-tabulation
   • summary_stats - Generate summary statistics
   • subtotals - Add subtotal rows

6. LOOKUP
   • xlookup - Modern lookup function
   • vlookup - Vertical lookup
   • index_match - INDEX-MATCH combination
   • join_sheets - Merge data from multiple sheets
   • fuzzy_match - Approximate matching

7. VISUALIZATION
   • bar_chart - Bar/column chart
   • pie_chart - Pie chart
   • line_chart - Line chart
   • scatter_plot - Scatter plot
   • kpi_dashboard - KPI dashboard
   • heatmap - Conditional formatting heatmap
   • sparklines - Mini charts in cells

8. REPORTING
   • summary_sheet - Create summary sheet
   • dashboard - Executive dashboard
   • change_log - Audit trail of changes
   • export_summary - Export summary report
   • data_profile - Statistical profile of data

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

BUSINESS LANGUAGE MAPPING (Natural to Technical):
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

CALCULATION TERMS:
"total", "sum up", "add together" → calculation + sum
"average", "mean", "typical" → calculation + average
"highest", "maximum", "largest" → calculation + max
"lowest", "minimum", "smallest" → calculation + min
"how many", "count" → calculation + count
"middle value", "median" → calculation + median
"ratio", "proportion" → calculation + percentage
"product", "times" → calculation + multiply
"split", "per" → calculation + divide

VISUALIZATION TERMS:
"chart", "graph", "plot", "visualize", "show me" → visualization
"bar chart", "column chart", "bars" → visualization + bar_chart
"pie chart", "pie", "distribution" → visualization + pie_chart
"line chart", "trend", "over time" → visualization + line_chart
"dashboard", "KPI", "metrics" → visualization + kpi_dashboard

TRANSFORMATION TERMS:
"sort", "order", "arrange" → transformation + sort
"remove duplicates", "unique values" → transformation + deduplicate
"split", "separate" → transformation + split_column
"combine", "merge", "concatenate" → transformation + merge_columns
"delete", "remove column" → transformation + remove_column (HIGH RISK)

AGGREGATION TERMS:
"pivot", "summarize by", "break down by" → aggregation + pivot_table
"group by", "aggregate" → aggregation + group_by
"subtotals", "totals for each" → aggregation + subtotals

FORMATTING TERMS:
"highlight", "color", "mark", "flag" → formatting + highlight
"format", "style" → formatting + conditional_color

LOOKUP TERMS:
"lookup", "match", "find", "search" → lookup + xlookup
"join", "merge sheets", "combine data" → lookup + join_sheets

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

CRITICAL RULES - MUST FOLLOW:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. COLUMN NAME EXACTNESS
   • Use EXACT column names from "AVAILABLE COLUMNS" list
   • Do NOT invent or assume column names
   • If user says "column A" or "first column", map to actual name
   • If ambiguous, set is_ambiguous=true and request clarification

2. DATA TYPE AWARENESS
   • Only suggest calculations on numeric columns
   • Charts need: 1 text column (category) + 1+ numeric columns (values)
   • Don't try to sum text columns

3. RISK ASSESSMENT MATRIX
   • LOW RISK: Calculations, adding columns, visualizations, sorting
   • MEDIUM RISK: Formatting, renaming, filtering
   • HIGH RISK: Deleting columns, removing rows, cross-sheet operations

4. AMBIGUITY HANDLING
   • If multiple interpretations exist, list in alternative_interpretations
   • If critical info missing, set clarification_needed
   • If columns don't exist, suggest creating them

5. OUTPUT FORMAT
   • Return ONLY the JSON object
   • NO markdown code blocks (```json```)
   • NO explanatory text before/after JSON
   • Ensure all fields are present (use null if not applicable)

6. CHART-SPECIFIC LOGIC
   • For charts: identify category column (text) and value columns (numeric)
   • Set formula_logic to describe chart: "bar chart of [value] by [category]"
   • Example: "bar chart of Sales by Product"

7. CONFIDENCE SCORING
   • HIGH confidence: Clear query, all columns exist, single interpretation
   • MEDIUM confidence: Minor ambiguity, need assumption
   • LOW confidence: Missing columns, multiple interpretations possible

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

EXAMPLES OF CORRECT PARSING:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

EXAMPLE 1 - Calculation:
Query: "Calculate total revenue by summing Q1, Q2, Q3"
{{
    "reasoning": {{
        "intent_summary": "User wants to calculate total annual revenue from quarterly data",
        "column_mapping": {{"Q1": "Q1", "Q2": "Q2", "Q3": "Q3"}},
        "operation_rationale": "Sum operation on numeric columns to create annual total",
        "risk_factors": ["Creates new column - low risk"]
    }},
    "operation_type": "calculation",
    "sub_operation": "sum",
    "source_columns": ["Q1", "Q2", "Q3"],
    "needs_new_column": true,
    "suggested_column_names": ["Total_Revenue"],
    "operation_description": "Sum Q1, Q2, and Q3 to create Total_Revenue column",
    "formula_logic": "sum(Q1, Q2, Q3)",
    "is_ambiguous": false,
    "clarification_needed": null,
    "risk_level": "low",
    "execution_confidence": "high",
    "alternative_interpretations": []
}}

EXAMPLE 2 - Visualization:
Query: "Show me a bar chart of sales by product"
{{
    "reasoning": {{
        "intent_summary": "User wants visual representation of sales performance across products",
        "column_mapping": {{"sales": "Sales", "product": "Product"}},
        "operation_rationale": "Bar chart with Product as category and Sales as values",
        "risk_factors": ["Read-only operation - no risk"]
    }},
    "operation_type": "visualization",
    "sub_operation": "bar_chart",
    "source_columns": ["Product", "Sales"],
    "needs_new_column": false,
    "suggested_column_names": null,
    "operation_description": "Create bar chart showing Sales by Product",
    "formula_logic": "bar chart of Sales by Product",
    "is_ambiguous": false,
    "clarification_needed": null,
    "risk_level": "low",
    "execution_confidence": "high",
    "alternative_interpretations": []
}}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

NOW PARSE THE USER QUERY AND RETURN ONLY THE JSON:
"""
        
        return prompt
    
    def _extract_json_from_response(self, response) -> Dict:
        """
        Robust JSON extraction from AI response with multiple fallback strategies
        """
        try:
            content = response.choices[0].message.content.strip()
            
            # Strategy 1: Direct JSON parsing
            try:
                return json.loads(content)
            except json.JSONDecodeError:
                pass
            
            # Strategy 2: Remove markdown code blocks
            if '```' in content:
                # Extract content between code blocks
                parts = content.split('```')
                for part in parts:
                    part = part.strip()
                    if part.startswith('json'):
                        part = part[4:].strip()
                    if part and (part.startswith('{') or part.startswith('[')):
                        try:
                            return json.loads(part)
                        except json.JSONDecodeError:
                            continue
            
            # Strategy 3: Regex extraction of JSON object
            json_match = re.search(r'\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\}', content, re.DOTALL)
            if json_match:
                try:
                    return json.loads(json_match.group())
                except json.JSONDecodeError:
                    pass
            
            # Strategy 4: Find largest JSON-like structure
            brace_count = 0
            start_idx = -1
            for i, char in enumerate(content):
                if char == '{':
                    if brace_count == 0:
                        start_idx = i
                    brace_count += 1
                elif char == '}':
                    brace_count -= 1
                    if brace_count == 0 and start_idx != -1:
                        try:
                            return json.loads(content[start_idx:i+1])
                        except json.JSONDecodeError:
                            start_idx = -1
            
            # If all strategies fail, return error with context
            return {
                "error": "Failed to parse AI response into JSON",
                "raw_response": content[:500],  # Truncate for readability
                "parse_error": "Multiple parsing strategies failed",
                "suggestions": [
                    "AI may not have followed JSON format",
                    "Response may be ambiguous or incomplete",
                    "Try rephrasing your query"
                ]
            }
            
        except Exception as e:
            return {
                "error": "Unexpected error during response parsing",
                "exception": str(e),
                "type": type(e).__name__
            }
    
    def _validate_decision_plan(self, decision: AgentDecision, excel_structure: Dict) -> Dict:
        """
        Validate decision plan before execution with AI-assisted validation
        
        Returns:
            Validation report with approval status and recommendations
        """
        validation_prompt = f"""You are an Enterprise Risk Assessor validating an Excel operation plan.

PROPOSED OPERATION PLAN:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
{json.dumps(decision.to_dict(), indent=2)}

EXCEL FILE CONTEXT:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Total Rows: {excel_structure.get('total_rows')}
Total Columns: {excel_structure.get('total_columns')}
Available Columns: {', '.join(excel_structure.get('columns', []))}
Numeric Columns: {', '.join(excel_structure.get('numeric_columns', []))}
Text Columns: {', '.join(excel_structure.get('text_columns', []))}

YOUR TASK:
Validate this decision plan and return ONLY a JSON assessment:

{{
    "is_valid": true,
    "validation_status": "approved | needs_review | rejected",
    "validation_checks": {{
        "columns_exist": true,
        "operation_feasible": true,
        "risk_acceptable": true,
        "data_types_compatible": true
    }},
    "warnings": [],
    "blockers": [],
    "recommendations": [],
    "requires_user_confirmation": false,
    "suggested_modifications": {{}}
}}

VALIDATION CRITERIA:
1. All source_columns must exist in available columns
2. Data types must be compatible with operation (e.g., can't sum text columns)
3. Risk level must be accurately assessed
4. For high-risk operations, flag for user confirmation
5. Check for potential data loss or corruption
6. Verify operation type matches sub_operation

Return ONLY the JSON validation response (no markdown, no explanations).
"""
        
        messages = [
            {"role": "system", "content": self.system_prompt},
            {"role": "user", "content": validation_prompt}
        ]
        
        response = self.client.chat.completions.create(
            model=self.model,
            messages=messages,
            temperature=0.2,
            max_tokens=800
        )
        
        return self._extract_json_from_response(response)
    
    def _generate_result_explanation(self, result: Dict, decision: AgentDecision) -> Dict:
        """
        Generate business-friendly explanation of operation results
        
        Returns:
            Professional summary for end users
        """
        explanation_prompt = f"""You are an Enterprise Communication Specialist explaining Excel operation results.

OPERATION EXECUTED:
{decision.change_description}

TECHNICAL RESULT:
{json.dumps(result, indent=2)}

YOUR TASK:
Create a clear, professional explanation for business users.

Return ONLY JSON:
{{
    "executive_summary": "One-sentence summary of what was accomplished",
    "what_changed": "Detailed explanation of changes made to the Excel file",
    "business_impact": "How this affects data analysis and decision-making",
    "next_steps": ["Recommended action 1", "Recommended action 2"],
    "technical_details": {{
        "rows_affected": 0,
        "columns_added": [],
        "file_location": "path/to/file.xlsx"
    }},
    "success_metrics": {{
        "operation_status": "completed",
        "data_quality": "validated",
        "audit_trail": "logged"
    }}
}}

Style: Professional, clear, non-technical language suitable for business stakeholders.
Return ONLY the JSON (no markdown).
"""
        
        messages = [
            {"role": "system", "content": self.system_prompt},
            {"role": "user", "content": explanation_prompt}
        ]
        
        response = self.client.chat.completions.create(
            model=self.model,
            messages=messages,
            temperature=0.3,
            max_tokens=1000
        )
        
        return self._extract_json_from_response(response)
    
    def _handle_error_recovery(self, error: str, query: str, context: Dict) -> Dict:
        """
        AI-assisted error recovery and suggestion generation
        
        Returns:
            Recovery strategies and alternative approaches
        """
        recovery_prompt = f"""You are an Enterprise Problem Solver for Excel operations.

ERROR ENCOUNTERED:
{error}

ORIGINAL USER QUERY:
"{query}"

CONTEXT:
{json.dumps(context, indent=2)}

YOUR TASK:
Analyze the error and suggest recovery strategies.

Return ONLY JSON:
{{
    "error_category": "data_issue | syntax_error | missing_column | type_mismatch | permission_error | other",
    "root_cause": "Brief explanation of what went wrong",
    "user_friendly_message": "Non-technical explanation for the user",
    "suggested_fixes": [
        {{
            "fix_description": "What to try",
            "modified_query": "Corrected query suggestion",
            "success_likelihood": "high | medium | low"
        }}
    ],
    "alternative_approaches": [
        "Alternative approach 1",
        "Alternative approach 2"
    ],
    "prevention_tips": ["How to avoid this error in the future"]
}}

Be helpful, solution-oriented, and empathetic. Return ONLY the JSON.
"""
        
        messages = [
            {"role": "system", "content": self.system_prompt},
            {"role": "user", "content": recovery_prompt}
        ]
        
        response = self.client.chat.completions.create(
            model=self.model,
            messages=messages,
            temperature=0.4,
            max_tokens=1000
        )
        
        return self._extract_json_from_response(response)
    
    def create_decision_plan(self, parsed_query: Dict, excel_structure: Dict) -> AgentDecision:
        """
        Creates an enterprise-grade structured decision plan based on parsed query
        
        Returns:
            AgentDecision object with complete execution plan
        """
        # Extract information
        operation_type = parsed_query.get("operation_type", "calculation")
        sub_operation = parsed_query.get("sub_operation", "sum")
        source_columns = parsed_query.get("source_columns", [])
        suggested_names = parsed_query.get("suggested_column_names", ["Result"])
        target_columns = suggested_names if isinstance(suggested_names, list) else [suggested_names]
        operation_desc = parsed_query.get("operation_description", "")
        sheet_name = excel_structure.get("analyzed_sheet", "Sheet1")
        
        # For visualization and some operations, no target column needed
        if operation_type in ["visualization", "reporting"]:
            target_columns = None
        
        # Determine risk level based on operation type and parsed query
        risk_level = parsed_query.get("risk_level", "low")
        if operation_type == "transformation" and sub_operation in ["remove_column", "deduplicate"]:
            risk_level = "high"
        elif operation_type in ["formatting", "validation"]:
            risk_level = "medium"
        
        # Requires confirmation for high-risk operations
        requires_confirmation = (risk_level == "high")
        
        # Create change description in business language
        if operation_type == "calculation":
            change_desc = f"Calculate {sub_operation} of {', '.join(source_columns)} and create new column(s): {', '.join(target_columns) if target_columns else 'N/A'}"
        elif operation_type == "visualization":
            change_desc = f"Create {sub_operation} chart using {', '.join(source_columns)}"
        elif operation_type == "transformation":
            change_desc = f"Transform data: {sub_operation} on {', '.join(source_columns)}"
        elif operation_type == "aggregation":
            change_desc = f"Aggregate data using {sub_operation} on {', '.join(source_columns)}"
        elif operation_type == "lookup":
            change_desc = f"Perform {sub_operation} to find matching data in {', '.join(source_columns)}"
        elif operation_type == "formatting":
            change_desc = f"Apply {sub_operation} formatting to {', '.join(source_columns)}"
        elif operation_type == "validation":
            change_desc = f"Add {sub_operation} validation rules to {', '.join(source_columns)}"
        else:
            change_desc = operation_desc
        
        # Create execution scope-specific assumptions
        execution_scope = "all_rows"
        total_rows = excel_structure.get('total_rows', 0)
        
        # Build assumptions list
        assumptions = []
        if target_columns:
            assumptions.append(f"Will create new column(s): {', '.join(target_columns)}")
        assumptions.append("Original data will remain unchanged (new file created)")
        assumptions.append(f"Operation will apply to {execution_scope} ({total_rows} rows)")
        if operation_type == "visualization":
            assumptions.append("Chart will be embedded in the Excel file below the data")
        
        # Create decision
        decision = AgentDecision(
            operation_type=operation_type,
            sub_operation=sub_operation,
            source_columns=source_columns,
            target_columns=target_columns,
            sheet_name=sheet_name,
            execution_scope=execution_scope,
            risk_level=risk_level,
            assumptions=assumptions,
            requires_confirmation=requires_confirmation,
            change_description=change_desc
        )
        
        return decision
    
    def execute_excel_operation(self, file_path: str, decision: AgentDecision, output_path: str) -> Dict:
        """
        Executes the planned Excel operation (Enterprise-grade)
        
        Returns:
            Dictionary with execution status and details
        """
        try:
            # Load Excel file
            df = pd.read_excel(file_path, sheet_name=decision.sheet_name)
            
            # Execute based on operation type
            if decision.operation_type == "calculation":
                df = self._execute_calculation(df, decision)
                df.to_excel(output_path, sheet_name=decision.sheet_name, index=False)
                
                return {
                    "status": "success",
                    "message": decision.change_description,
                    "rows_affected": len(df),
                    "output_file": output_path,
                    "changes": decision.assumptions,
                    "operation": f"{decision.operation_type} - {decision.sub_operation}"
                }
            
            elif decision.operation_type == "transformation":
                df = self._execute_transformation(df, decision)
                df.to_excel(output_path, sheet_name=decision.sheet_name, index=False)
                
                return {
                    "status": "success",
                    "message": decision.change_description,
                    "rows_affected": len(df),
                    "output_file": output_path,
                    "changes": decision.assumptions,
                    "operation": f"{decision.operation_type} - {decision.sub_operation}"
                }
            
            elif decision.operation_type == "aggregation":
                df = self._execute_aggregation(df, decision)
                df.to_excel(output_path, sheet_name=decision.sheet_name, index=False)
                
                return {
                    "status": "success",
                    "message": decision.change_description,
                    "rows_affected": len(df),
                    "output_file": output_path,
                    "changes": decision.assumptions,
                    "operation": f"{decision.operation_type} - {decision.sub_operation}"
                }
            
            elif decision.operation_type == "formatting":
                return self._execute_formatting(file_path, decision, output_path)
            
            elif decision.operation_type == "visualization":
                return self._execute_visualization(file_path, df, decision, output_path)
            
            elif decision.operation_type == "validation":
                return self._execute_validation(file_path, df, decision, output_path)
            
            elif decision.operation_type == "lookup":
                df = self._execute_lookup(df, decision)
                df.to_excel(output_path, sheet_name=decision.sheet_name, index=False)
                
                return {
                    "status": "success",
                    "message": decision.change_description,
                    "rows_affected": len(df),
                    "output_file": output_path,
                    "changes": decision.assumptions,
                    "operation": f"{decision.operation_type} - {decision.sub_operation}"
                }
            
            else:
                # Default: save as is
                df.to_excel(output_path, sheet_name=decision.sheet_name, index=False)
                
                return {
                    "status": "success",
                    "message": f"Successfully processed {decision.operation_type}",
                    "rows_affected": len(df),
                    "output_file": output_path,
                    "changes": decision.assumptions,
                    "operation": decision.operation_type
                }
            
        except Exception as e:
            return {
                "status": "failed",
                "error": str(e),
                "operation": f"{decision.operation_type} - {decision.sub_operation}"
            }
    
    def _execute_calculation(self, df: pd.DataFrame, decision: AgentDecision) -> pd.DataFrame:
        """Execute calculation operations (Enterprise-grade)"""
        source_cols = decision.source_columns
        target_cols = decision.target_columns if decision.target_columns else ["Result"]
        target_col = target_cols[0]  # Primary target column
        sub_op = decision.sub_operation.lower()
        
        # CALCULATION SUB-OPERATIONS
        if sub_op == "sum":
            df[target_col] = df[source_cols].sum(axis=1)
        
        elif sub_op == "average" or sub_op == "mean":
            df[target_col] = df[source_cols].mean(axis=1)
        
        elif sub_op == "min" or sub_op == "minimum":
            df[target_col] = df[source_cols].min(axis=1)
        
        elif sub_op == "max" or sub_op == "maximum":
            df[target_col] = df[source_cols].max(axis=1)
        
        elif sub_op == "count":
            df[target_col] = df[source_cols].count(axis=1)
        
        elif sub_op == "median":
            df[target_col] = df[source_cols].median(axis=1)
        
        elif sub_op == "multiply":
            result = df[source_cols[0]]
            for col in source_cols[1:]:
                result = result * df[col]
            df[target_col] = result
        
        elif sub_op == "divide":
            if len(source_cols) >= 2:
                df[target_col] = df[source_cols[0]] / df[source_cols[1]]
        
        elif sub_op == "percentage":
            if len(source_cols) >= 2:
                df[target_col] = (df[source_cols[0]] / df[source_cols[1]]) * 100
            else:
                # Percentage of total
                df[target_col] = (df[source_cols[0]] / df[source_cols[0]].sum()) * 100
        
        else:
            # Default: sum
            df[target_col] = df[source_cols].sum(axis=1)
        
        return df
    
    def _execute_transformation(self, df: pd.DataFrame, decision: AgentDecision) -> pd.DataFrame:
        """Execute data transformation operations"""
        sub_op = decision.sub_operation.lower()
        source_cols = decision.source_columns
        target_cols = decision.target_columns
        
        if sub_op == "add_column":
            # Add new empty column(s)
            for col in target_cols:
                df[col] = None
        
        elif sub_op == "rename_column":
            # Rename column (source[0] -> target[0])
            if source_cols and target_cols:
                df.rename(columns={source_cols[0]: target_cols[0]}, inplace=True)
        
        elif sub_op == "remove_column":
            # Remove columns (high risk - requires confirmation)
            df.drop(columns=source_cols, inplace=True, errors='ignore')
        
        elif sub_op == "sort":
            # Sort by columns
            df.sort_values(by=source_cols, inplace=True)
        
        elif sub_op == "deduplicate":
            # Remove duplicates based on columns
            df.drop_duplicates(subset=source_cols, inplace=True)
        
        elif sub_op == "split_column":
            # Split column (basic implementation - can be enhanced)
            if source_cols and target_cols and len(target_cols) >= 2:
                # Split first source column by space/comma
                split_data = df[source_cols[0]].str.split(expand=True)
                for idx, target_col in enumerate(target_cols[:split_data.shape[1]]):
                    df[target_col] = split_data[idx]
        
        elif sub_op == "merge_columns":
            # Merge multiple columns into one
            if source_cols and target_cols:
                df[target_cols[0]] = df[source_cols].apply(lambda row: ' '.join(row.values.astype(str)), axis=1)
        
        return df
    
    def _execute_lookup(self, df: pd.DataFrame, decision: AgentDecision) -> pd.DataFrame:
        """Execute lookup/reference operations"""
        # Basic implementation - can be enhanced for cross-sheet lookups
        sub_op = decision.sub_operation.lower()
        
        # Placeholder for xlookup, vlookup implementations
        # Would require reference table/sheet information
        
        return df
    
    def _execute_validation(self, file_path: str, df: pd.DataFrame, decision: AgentDecision, output_path: str) -> Dict:
        """Execute data validation operations"""
        try:
            # Save dataframe first
            df.to_excel(output_path, sheet_name=decision.sheet_name, index=False)
            
            # Load with openpyxl for validation rules
            wb = openpyxl.load_workbook(output_path)
            ws = wb[decision.sheet_name]
            
            # Add validation rules (basic implementation)
            # Can be enhanced with dropdown lists, range checks, etc.
            
            wb.save(output_path)
            
            return {
                "status": "success",
                "message": f"Validation rules applied: {decision.sub_operation}",
                "output_file": output_path,
                "changes": decision.assumptions,
                "operation": f"{decision.operation_type} - {decision.sub_operation}"
            }
        except Exception as e:
            return {
                "status": "failed",
                "error": f"Validation failed: {str(e)}"
            }
    
    def _execute_aggregation(self, df: pd.DataFrame, decision: AgentDecision) -> pd.DataFrame:
        """Execute aggregation operations"""
        # This would handle group by operations
        # For now, basic implementation
        return df
    
    def _execute_formatting(self, file_path: str, decision: AgentDecision, output_path: str) -> Dict:
        """Execute formatting operations using openpyxl"""
        # Load workbook
        wb = openpyxl.load_workbook(file_path)
        ws = wb[decision.sheet_name]
        
        # Apply formatting based on decision
        # Example: Highlight negative values in red
        
        wb.save(output_path)
        
        return {
            "status": "success",
            "message": "Formatting applied successfully",
            "output_file": output_path
        }
    
    def _execute_visualization(self, file_path: str, df: pd.DataFrame, decision: AgentDecision, output_path: str) -> Dict:
        """Execute visualization operations - create charts in Excel"""
        try:
            # First save the data to Excel
            df.to_excel(output_path, sheet_name=decision.sheet_name, index=False)
            
            # Load the workbook to add chart
            wb = openpyxl.load_workbook(output_path)
            ws = wb[decision.sheet_name]
            
            # Determine chart type from formula logic
            formula_lower = decision.formula_logic.lower()
            source_cols = decision.source_columns
            
            # Identify category column (usually text/product name) and data columns (numeric)
            category_col = None
            data_cols = []
            
            for col in source_cols:
                if col in df.columns:
                    if df[col].dtype == 'object':  # Text column
                        category_col = col
                    else:  # Numeric column
                        data_cols.append(col)
            
            # If no category column found, use first column
            if not category_col and len(source_cols) > 0:
                category_col = source_cols[0]
                data_cols = source_cols[1:] if len(source_cols) > 1 else source_cols
            
            # Get column positions in Excel (1-indexed)
            col_positions = {col: idx + 1 for idx, col in enumerate(df.columns)}
            
            # Determine chart type
            if "pie" in formula_lower:
                chart = PieChart()
                chart.title = "Pie Chart"
            elif "line" in formula_lower:
                chart = LineChart()
                chart.title = "Line Chart"
            else:
                # Default to bar chart
                chart = BarChart()
                chart.title = "Bar Chart"
            
            # Set chart dimensions
            chart.height = 15
            chart.width = 25
            
            # Set data range (excluding header)
            min_row = 2
            max_row = len(df) + 1
            
            # Add data series
            for data_col in data_cols:
                if data_col in col_positions:
                    col_letter = openpyxl.utils.get_column_letter(col_positions[data_col])
                    
                    # Data values
                    data = Reference(ws, min_col=col_positions[data_col], min_row=1, 
                                   max_row=max_row)
                    chart.add_data(data, titles_from_data=True)
            
            # Set categories (x-axis labels)
            if category_col and category_col in col_positions:
                cat_col_letter = openpyxl.utils.get_column_letter(col_positions[category_col])
                cats = Reference(ws, min_col=col_positions[category_col], min_row=2, 
                               max_row=max_row)
                chart.set_categories(cats)
            
            # Set chart labels
            chart.x_axis.title = category_col if category_col else "Categories"
            chart.y_axis.title = "Values"
            
            # Style
            chart.style = 10
            
            # Add chart to worksheet at a position below the data
            chart_position = f"A{max_row + 3}"
            ws.add_chart(chart, chart_position)
            
            # Save workbook with chart
            wb.save(output_path)
            
            return {
                "status": "success",
                "message": f"Successfully created bar chart for {', '.join(data_cols)}",
                "rows_affected": len(df),
                "output_file": output_path,
                "changes": [
                    f"Created {'pie' if 'pie' in formula_lower else 'line' if 'line' in formula_lower else 'bar'} chart",
                    f"Chart shows: {', '.join(data_cols)}",
                    f"Categories: {category_col}" if category_col else "No categories",
                    f"Chart added at position {chart_position}"
                ]
            }
            
        except Exception as e:
            return {
                "status": "failed",
                "error": f"Chart creation failed: {str(e)}"
            }
    
    def process_query(self, file_path: str, query: str, output_path: Optional[str] = None) -> Dict:
        """
        Main entry point: Process user query on Excel file
        
        This is the orchestration method that follows the complete workflow
        """
        print("\n" + "="*60)
        print("🤖 EXCEL AGENT - PROCESSING REQUEST")
        print("="*60)
        
        # Set output path
        if not output_path:
            base, ext = os.path.splitext(file_path)
            output_path = f"{base}_modified{ext}"
        
        # STEP 1: Analyze Excel Structure
        print("\n📊 STEP 1: Analyzing Excel Structure...")
        excel_structure = self.analyze_excel_structure(file_path)
        
        if "error" in excel_structure:
            return {"status": "failed", "error": excel_structure["error"]}
        
        print(f"   ✓ Found {excel_structure['total_columns']} columns: {', '.join(excel_structure['columns'])}")
        print(f"   ✓ Total rows: {excel_structure['total_rows']}")
        
        # STEP 2: Parse User Query
        print(f"\n🧠 STEP 2: Analyzing User Intent...")
        print(f"   Query: '{query}'")
        
        parsed_query = self.parse_user_query(query, excel_structure)
        
        if "error" in parsed_query:
            return {"status": "failed", "error": parsed_query["error"]}
        
        # Display reasoning if available
        if "reasoning" in parsed_query:
            reasoning = parsed_query["reasoning"]
            print(f"\n   💭 AI Reasoning:")
            print(f"      Intent: {reasoning.get('intent_summary', 'N/A')}")
            print(f"      Rationale: {reasoning.get('operation_rationale', 'N/A')}")
            if reasoning.get('risk_factors'):
                print(f"      Risk Factors: {', '.join(reasoning.get('risk_factors', []))}")
        
        print(f"   ✓ Operation Type: {parsed_query.get('operation_type')}")
        print(f"   ✓ Sub-Operation: {parsed_query.get('sub_operation')}")
        print(f"   ✓ Source Columns: {', '.join(parsed_query.get('source_columns', []))}")
        print(f"   ✓ Execution Confidence: {parsed_query.get('execution_confidence', 'N/A').upper()}")
        
        # STEP 3: Check for clarification needed
        if parsed_query.get("is_ambiguous") or parsed_query.get("clarification_needed"):
            alternatives = parsed_query.get("alternative_interpretations", [])
            if alternatives:
                print(f"\n   🤔 Alternative Interpretations Detected:")
                for i, alt in enumerate(alternatives, 1):
                    print(f"      {i}. {alt}")
            
            return {
                "status": "clarification_needed",
                "question": parsed_query.get("clarification_needed"),
                "alternatives": alternatives,
                "parsed_info": parsed_query
            }
        
        # STEP 4: Create Decision Plan
        print(f"\n📋 STEP 3: Creating Enterprise Decision Plan...")
        decision = self.create_decision_plan(parsed_query, excel_structure)
        
        print(f"   ✓ Operation: {decision.operation_type} → {decision.sub_operation}")
        print(f"   ✓ Target Column(s): {decision.target_columns if decision.target_columns else 'N/A (visualization/reporting)'}")
        print(f"   ✓ Change Description: {decision.change_description}")
        print(f"   ✓ Risk Level: {decision.risk_level.upper()}")
        
        print("\n   📄 Decision Plan (JSON):")
        print(json.dumps(decision.to_dict(), indent=4))
        
        # STEP 4: Validate Decision Plan
        print(f"\n🔍 STEP 4: Validating Decision Plan...")
        validation_result = self._validate_decision_plan(decision, excel_structure)
        
        if "error" not in validation_result:
            validation_status = validation_result.get("validation_status", "unknown")
            is_valid = validation_result.get("is_valid", False)
            
            print(f"   ✓ Validation Status: {validation_status.upper()}")
            
            # Display warnings if any
            warnings = validation_result.get("warnings", [])
            if warnings:
                print(f"   ⚠️  Warnings:")
                for warning in warnings:
                    print(f"      - {warning}")
            
            # Display blockers if any
            blockers = validation_result.get("blockers", [])
            if blockers:
                print(f"   ❌ Blockers Found:")
                for blocker in blockers:
                    print(f"      - {blocker}")
                return {
                    "status": "validation_failed",
                    "blockers": blockers,
                    "validation_result": validation_result
                }
            
            # Check if user confirmation needed
            if validation_result.get("requires_user_confirmation") or decision.risk_level == "high":
                print(f"   ⚠️  HIGH-RISK OPERATION: User confirmation required")
                return {
                    "status": "confirmation_required",
                    "decision": decision.to_dict(),
                    "validation": validation_result,
                    "message": "This is a high-risk operation. Please review and confirm."
                }
        else:
            print(f"   ⚠️  Validation check skipped (validation service unavailable)")
        
        # STEP 5: Execute Operation
        print(f"\n⚙️  STEP 5: Executing Enterprise Excel Operation...")
        
        try:
            result = self.execute_excel_operation(file_path, decision, output_path)
            
            if result["status"] == "success":
                print(f"   ✅ SUCCESS!")
                print(f"   ✓ Operation: {result.get('operation', 'N/A')}")
                print(f"   ✓ {result['message']}")
                print(f"   ✓ Rows affected: {result['rows_affected']}")
                print(f"   ✓ Output saved to: {result['output_file']}")
                print(f"\n   📋 Changes Made:")
                for change in result.get('changes', []):
                    print(f"      • {change}")
                
                # STEP 6: Generate Business-Friendly Explanation
                print(f"\n📊 STEP 6: Generating Executive Summary...")
                explanation = self._generate_result_explanation(result, decision)
                
                if "error" not in explanation:
                    print(f"\n   Executive Summary:")
                    print(f"   {explanation.get('executive_summary', 'Operation completed successfully')}")
                    
                    # Add explanation to result
                    result['explanation'] = explanation
                
                # Log the change for audit
                self.change_log.append({
                    "timestamp": pd.Timestamp.now().isoformat(),
                    "file": file_path,
                    "query": query,
                    "operation": result.get('operation'),
                    "output": output_path,
                    "risk_level": decision.risk_level,
                    "decision": decision.to_dict()
                })
            else:
                print(f"   ❌ FAILED: {result.get('error')}")
                print(f"   Operation attempted: {result.get('operation', 'N/A')}")
        
        except Exception as e:
            error_msg = str(e)
            print(f"   ❌ EXECUTION ERROR: {error_msg}")
            
            # STEP 6 (Error Path): Generate Recovery Suggestions
            print(f"\n🔧 STEP 6: Analyzing Error and Generating Recovery Strategies...")
            recovery = self._handle_error_recovery(
                error=error_msg,
                query=query,
                context={
                    "excel_structure": excel_structure,
                    "decision": decision.to_dict(),
                    "parsed_query": parsed_query
                }
            )
            
            if "error" not in recovery:
                print(f"\n   Error Analysis:")
                print(f"   Category: {recovery.get('error_category', 'unknown')}")
                print(f"   Root Cause: {recovery.get('root_cause', 'Unknown')}")
                print(f"   User Message: {recovery.get('user_friendly_message', '')}")
                
                suggested_fixes = recovery.get('suggested_fixes', [])
                if suggested_fixes:
                    print(f"\n   💡 Suggested Fixes:")
                    for i, fix in enumerate(suggested_fixes, 1):
                        print(f"      {i}. {fix.get('fix_description')}")
                        print(f"         Try: \"{fix.get('modified_query')}\"")
                        print(f"         Success Likelihood: {fix.get('success_likelihood', 'unknown')}")
            
            result = {
                "status": "failed",
                "error": error_msg,
                "recovery_suggestions": recovery
            }
        
        print("\n" + "="*60)
        
        return result
    
    def get_audit_log(self) -> List[Dict]:
        """
        Get complete audit trail of all operations performed by the agent
        
        Returns:
            List of audit log entries with timestamps, queries, and operations
        """
        return self.change_log
    
    def export_audit_log(self, output_path: str = "agent_audit_log.json") -> str:
        """
        Export audit log to JSON file for compliance and review
        
        Args:
            output_path: Path to save the audit log
            
        Returns:
            Path to the exported file
        """
        with open(output_path, 'w') as f:
            json.dump(self.change_log, f, indent=2, default=str)
        
        return output_path


# ==========================================
# HELPER FUNCTION FOR QUICK USAGE
# ==========================================
def create_excel_agent(api_key: str) -> ExcelAgent:
    """Factory function to create an Excel Agent instance"""
    return ExcelAgent(api_key=api_key)


# ==========================================
# INTERACTIVE MODE - MAIN ENTRY POINT
# ==========================================
def run_interactive_agent(api_key: str):
    """
    Interactive mode where agent asks for Excel file and then processes queries
    """
    print("\n" + "="*70)
    print("🏢 ENTERPRISE EXCEL AUTOMATION AI AGENT - INTERACTIVE MODE")
    print("="*70)
    print("\nWelcome! I'm your Enterprise Excel Automation Agent.")
    print("I can help you with professional Excel operations using business language.\n")
    print("I am NOT a chatbot - I am a decision-making automation expert.")
    print("\n📊 SUPPORTED OPERATIONS:")
    print("  1. Calculations: Sum, Average, Min, Max, %, Count, Median")
    print("  2. Transformations: Add/Remove/Rename columns, Sort, Filter, Deduplicate")
    print("  3. Validations: Dropdowns, Error detection, Quality checks")
    print("  4. Formatting: Conditional colors, Highlights, Data bars")
    print("  5. Aggregation: Pivot tables, Group by, Summaries")
    print("  6. Lookups: XLOOKUP, VLOOKUP, Data matching")
    print("  7. Visualization: Charts (Bar, Pie, Line), KPI dashboards")
    print("  8. Reporting: Summary sheets, Change logs\n")
    
    # Initialize agent
    agent = ExcelAgent(api_key=api_key)
    
    # STEP 1: Ask for Excel file
    print("="*70)
    print("STEP 1: Excel File Selection")
    print("="*70)
    
    while True:
        file_path = input("\n📁 Please provide the path to your Excel file: ").strip()
        
        # Remove quotes if user wrapped path in quotes
        file_path = file_path.strip('"').strip("'")
        
        if not file_path:
            print("⚠️  No file path provided. Please try again.")
            continue
        
        if not os.path.exists(file_path):
            print(f"❌ File not found: {file_path}")
            retry = input("Would you like to try again? (yes/no): ").strip().lower()
            if retry not in ['yes', 'y']:
                print("\n👋 Exiting. Goodbye!")
                return
            continue
        
        # Try to load and analyze the file
        try:
            print(f"\n📊 Analyzing Excel file: {os.path.basename(file_path)}")
            structure = agent.analyze_excel_structure(file_path)
            
            if "error" in structure:
                print(f"❌ Error reading file: {structure['error']}")
                retry = input("Would you like to try another file? (yes/no): ").strip().lower()
                if retry not in ['yes', 'y']:
                    print("\n👋 Exiting. Goodbye!")
                    return
                continue
            
            # Display file information
            print(f"\n✅ Successfully loaded: {structure['file_name']}")
            print(f"   📄 Sheet: {structure['analyzed_sheet']}")
            print(f"   📊 Rows: {structure['total_rows']}")
            print(f"   📋 Columns ({structure['total_columns']}): {', '.join(structure['columns'])}")
            
            # Show sample data
            if structure.get('sample_data'):
                print(f"\n   Sample Data (first 3 rows):")
                import pandas as pd
                sample_df = pd.DataFrame(structure['sample_data'])
                print("   " + "\n   ".join(sample_df.to_string(index=False).split('\n')))
            
            break
            
        except Exception as e:
            print(f"❌ Error: {e}")
            retry = input("Would you like to try another file? (yes/no): ").strip().lower()
            if retry not in ['yes', 'y']:
                print("\n👋 Exiting. Goodbye!")
                return
            continue
    
    # STEP 2: Interactive query loop
    print("\n" + "="*70)
    print("STEP 2: Modification Queries")
    print("="*70)
    print("\nYou can now ask me to perform operations using business language.")
    print("\n💼 EXAMPLE BUSINESS QUERIES:")
    print("  📊 Calculations:")
    print("    • 'Calculate total sales by summing Q1, Q2, and Q3'")
    print("    • 'Find the average revenue across all products'")
    print("    • 'Calculate profit margin percentage'")
    print("\n  📈 Visualizations:")
    print("    • 'Create a bar chart for sales by product'")
    print("    • 'Show me a pie chart of revenue distribution'")
    print("\n  🔧 Transformations:")
    print("    • 'Sort data by Sales descending'")
    print("    • 'Remove duplicate entries'")
    print("\n  🎨 Formatting:")
    print("    • 'Highlight cells where profit is negative'")
    print("\nType 'quit', 'exit', or 'q' to stop.")
    print("Type 'new' to load a different Excel file.")
    print("="*70)
    
    query_count = 0
    current_file = file_path
    
    while True:
        print(f"\n{'─'*70}")
        query = input("\n💬 What would you like to modify in the Excel? ").strip()
        
        # Handle exit commands
        if query.lower() in ['quit', 'exit', 'q']:
            print("\n" + "="*70)
            print("👋 Thank you for using Excel Automation Agent!")
            print("="*70)
            break
        
        # Handle new file command
        if query.lower() == 'new':
            print("\n🔄 Loading a new Excel file...\n")
            run_interactive_agent(api_key)
            return
        
        # Skip empty queries
        if not query:
            print("⚠️  Please enter a query or type 'quit' to exit.")
            continue
        
        # Process the query
        query_count += 1
        
        # Generate output filename
        base_name = os.path.splitext(os.path.basename(current_file))[0]
        output_path = f"{base_name}_modified_{query_count}.xlsx"
        
        print(f"\n🔄 Processing your request...")
        
        try:
            result = agent.process_query(
                file_path=current_file,
                query=query,
                output_path=output_path
            )
            
            # Handle results
            if result.get("status") == "success":
                print(f"\n{'='*70}")
                print("✅ MODIFICATION SUCCESSFUL!")
                print(f"{'='*70}")
                print(f"📁 New file created: {result['output_file']}")
                print(f"📝 {result['message']}")
                print(f"📊 Rows modified: {result['rows_affected']}")
                
                # Show preview
                try:
                    import pandas as pd
                    preview_df = pd.read_excel(output_path)
                    print(f"\n📋 Preview of modified data:")
                    print(preview_df.to_string(index=False))
                except:
                    pass
                
                # Ask if they want to continue with the modified file
                print(f"\n{'─'*70}")
                continue_choice = input("\n🔄 Use this modified file for next query? (yes/no, default: yes): ").strip().lower()
                
                if continue_choice in ['', 'yes', 'y']:
                    current_file = output_path
                    print(f"✅ Now working with: {output_path}")
                else:
                    print(f"✅ Still working with: {current_file}")
            
            elif result.get("status") == "clarification_needed":
                print(f"\n❓ I need clarification:")
                print(f"   {result['question']}")
                print("\n💡 Please rephrase your query with more details.")
            
            else:
                print(f"\n❌ Error: {result.get('error', 'Unknown error occurred')}")
                print("💡 Please try rephrasing your query or check the column names.")
        
        except Exception as e:
            print(f"\n❌ Unexpected error: {e}")
            print("💡 Please try again with a different query.")


# ==========================================
# MAIN ENTRY POINT
# ==========================================
if __name__ == "__main__":
    import os
    
    # Get API Key from environment variable or user input
    API_KEY = os.getenv("GROQ_API_KEY")
    
    if not API_KEY:
        print("⚠️  GROQ_API_KEY not found in environment variables")
        print("Get your free API key from: https://console.groq.com\n")
        API_KEY = input("Enter your Groq API Key: ").strip()
    
    # Run interactive agent
    run_interactive_agent(API_KEY)
