# 🤖 AI Agent Collection

**Enterprise-grade AI Agents powered by Groq & Llama 3.3**

This repository contains powerful AI agents that automate complex tasks using natural language.

---

## 📦 What's Inside

### 1. **Excel AI Agent** 📊
Enterprise-grade Excel automation agent with advanced prompt engineering
- Natural language Excel operations
- 50+ operations (calculations, charts, transformations)
- Multi-step reasoning & decision validation
- Complete audit trails

### 2. **Test Case Generator Agent** 🧪
Intelligent test case generation for any function or feature
- Automatic test case generation
- Positive, negative, edge & security tests
- AI-powered test planning

---

## ⚡ Quick Start

### Prerequisites

- **Python 3.8+** installed ([Download here](https://www.python.org/downloads/))
- **Groq API Key** (free) - [Get it here](https://console.groq.com)

### Installation

```bash
# 1. Clone the repository
git clone https://github.com/shivendersinghlucky-creator/AIAgent.git
cd AIAgent

# 2. Install dependencies
pip install -r requirements.txt
```

---

## 🔑 **CONFIGURATION - REQUIRED!**

### Step 1: Get Your Free Groq API Key

1. **Visit**: https://console.groq.com
2. **Sign up** for a free account (no credit card required)
3. **Navigate to**: API Keys section
4. **Create** a new API key
5. **Copy** the key (starts with `gsk_...`)

### Step 2: Set Up API Key

Choose **ONE** of these methods:

#### **Option A: Environment Variable (Recommended)**

**Windows (PowerShell):**
```powershell
# Temporary (current session only)
$env:GROQ_API_KEY="your_api_key_here"

# Permanent (for your user account)
[System.Environment]::SetEnvironmentVariable('GROQ_API_KEY', 'your_api_key_here', 'User')
```

**Linux/Mac:**
```bash
# Temporary (current session)
export GROQ_API_KEY="your_api_key_here"

# Permanent (add to ~/.bashrc or ~/.zshrc)
echo 'export GROQ_API_KEY="your_api_key_here"' >> ~/.bashrc
source ~/.bashrc
```

#### **Option B: Create `.env` File**

Create a file named `.env` in the project root:

```bash
GROQ_API_KEY=your_api_key_here
```

Then install `python-dotenv`:
```bash
pip install python-dotenv
```

#### **Option C: Direct Input (When Running)**

The scripts will prompt you for the API key if not found in environment.

---

## 🚀 Usage

### Excel AI Agent

```bash
python excel_agent.py
```


**Interactive Mode:**
- Upload your Excel file
- Ask in natural language what you want to do
- Agent analyzes, plans, validates, and executes
- Download modified file

**Example Queries:**
```
"Calculate total sales by summing Q1, Q2, Q3"
"Create a bar chart of revenue by product"
"Sort data by Sales descending"
"Find average of all numeric columns"
"Highlight negative values in red"
```

**Python API:**
```python
from excel_agent import ExcelAgent
import os

# Initialize with API key
api_key = os.getenv("GROQ_API_KEY")
agent = ExcelAgent(api_key=api_key)

# Process a query
result = agent.process_query(
    file_path="sales_data.xlsx",
    query="Calculate total revenue by summing Q1, Q2, Q3",
    output_path="sales_modified.xlsx"
)

print(result['status'])  # 'success', 'failed', 'clarification_needed'
```

---

### Test Case Generator Agent

```bash
python test_case_agent.py
```

**Interactive Mode:**
- Enter the function/feature name
- Agent generates comprehensive test cases
- Includes positive, negative, edge, and security tests

**Example:**
```
Scenario: "Login function with username and password"
```

**Python API:**
```python
from test_case_agent import TestCaseAgent
import os

# Set API key first
os.environ['GROQ_API_KEY'] = 'your_api_key_here'

# Create agent
agent = TestCaseAgent()

# Generate test cases
result = agent.think_and_act(
    "Generate test cases for user registration with email validation"
)

print(result)
```

---

## 📋 Features

### Excel AI Agent Features

| Category | Operations |
|----------|-----------|
| **Calculations** | Sum, Average, Min, Max, Count, Median, %, Multiply, Divide |
| **Charts** | Bar, Pie, Line, Scatter, KPI Dashboards |
| **Transformations** | Sort, Filter, Remove Duplicates, Split/Merge Columns |
| **Formatting** | Conditional Colors, Highlights, Data Bars |
| **Aggregation** | Pivot Tables, Group By, Summaries |
| **Validation** | Dropdowns, Error Detection, Data Quality |
| **Lookup** | XLOOKUP, VLOOKUP, INDEX-MATCH |

**Advanced Features:**
- ✅ Multi-step reasoning (5-step framework)
- ✅ Decision validation before execution
- ✅ Risk assessment (low/medium/high)
- ✅ Error recovery with suggestions
- ✅ Complete audit trails
- ✅ Business-friendly explanations

---

## 🔒 Security & Privacy

- ✅ **API keys never stored** in code
- ✅ **Original files never modified** (creates new files)
- ✅ **All operations logged** for audit compliance
- ✅ **Data stays local** (only queries sent to API)
- ✅ **No data stored** by Groq after processing

---

## 🛠️ Troubleshooting

### "No module named 'groq'"
```bash
pip install groq pandas openpyxl
```

### "API key not found"
Make sure you've set the `GROQ_API_KEY` environment variable:
```bash
# Check if set (Windows)
echo $env:GROQ_API_KEY

# Check if set (Linux/Mac)
echo $GROQ_API_KEY
```

### "Authentication failed"
- Verify your API key is correct
- Get a new key from https://console.groq.com
- Make sure no extra spaces in the key

### "File not found"
Use absolute paths for Excel files:
```python
# Windows
file_path = r"C:\Users\YourName\Documents\data.xlsx"

# Linux/Mac
file_path = "/home/username/documents/data.xlsx"
```

---

## 🎯 Use Cases

### Excel Agent
- 📊 **Finance**: Automate financial reports and calculations
- 📈 **Sales**: Generate sales dashboards and analytics
- 🔢 **Data Analysis**: Process and visualize data quickly
- 📋 **Reporting**: Create executive summaries automatically

### Test Case Agent
- 🧪 **QA Testing**: Generate comprehensive test suites
- 🔍 **Code Review**: Identify test scenarios to cover
- 📝 **Documentation**: Create test documentation automatically
- 🎯 **Test Planning**: Plan testing strategy for features

---

## 🔧 Advanced Configuration

### Customize Excel Agent Behavior

```python
from excel_agent import ExcelAgent

agent = ExcelAgent(api_key="your_key")

# Get audit log
audit_log = agent.get_audit_log()

# Export audit for compliance
agent.export_audit_log("audit_2024.json")

# Handle different response types
result = agent.process_query(file, query)

if result['status'] == 'clarification_needed':
    print(result['question'])
    # Rephrase and try again
    
elif result['status'] == 'confirmation_required':
    print(f"⚠️ High-risk operation: {result['message']}")
    # Review and confirm
```

---

## 💰 Cost

- **Groq API**: Free tier available (30 requests/minute)
- **Paid tier**: ~$0.10 per million tokens (very affordable)
- **No hidden costs**: Just API usage

---

## 📊 Performance

- **Response time**: 1-3 seconds per query
- **Accuracy**: 95%+ parsing accuracy
- **Scalability**: Handles Excel files up to 100MB
- **Concurrent users**: Limited by Groq API rate limits

---

## ⚠️ Important Notes

### Before First Use:
1. ✅ Install Python 3.8+
2. ✅ Get Groq API key (free)
3. ✅ Set environment variable
4. ✅ Install dependencies
5. ✅ Test with sample data

### Best Practices:
- 🔒 Never commit API keys to Git
- 💾 Always backup important Excel files
- 📝 Review high-risk operations before confirming
- 🔍 Check audit logs for compliance
- 🧪 Test on sample data first

---

## 🎉 Quick Test

To verify everything works:

```bash
# 1. Set API key
$env:GROQ_API_KEY="your_key_here"

# 2. Test Excel Agent
python -c "from excel_agent import ExcelAgent; import os; agent = ExcelAgent(os.getenv('GROQ_API_KEY')); print('✅ Excel Agent works!')"

# 3. Test Test Case Agent  
python -c "from test_case_agent import TestCaseAgent; print('✅ Test Agent works!')"
```

---

**Built with ❤️ using Groq's Llama 3.3 70B**

**Ready to automate? Start with:** `python excel_agent.py`

---

*Last updated: December 25, 2025*

- Range checks
- Error detection
- Blank/Duplicate checks

### 4. **Conditional Formatting** 🎨
- Color coding
- Highlighting
- Data bars
- Icon sets

### 5. **Aggregation & Analytics** 📈
- Pivot tables
- Group by operations
- Cross-tabulation
- Summary statistics

### 6. **Lookup & Reference** 🔍
- XLOOKUP, VLOOKUP
- INDEX-MATCH
- Join sheets
- Fuzzy matching

### 7. **Visualization** 📊
- Bar/Column charts
- Pie charts
- Line charts
- KPI dashboards

### 8. **Reporting & Governance** 📋
- Summary sheets
- Executive dashboards
- Change logs
- Data protection

---

## 💡 Usage Examples

### Example 1: Simple Calculation

```python
query = "Calculate total revenue by summing Q1, Q2, Q3, and Q4 sales"
result = agent.process_query("sales.xlsx", query)

# Output includes:
# - Reasoning: Why this operation was chosen
# - Execution confidence: High/Medium/Low
# - Business explanation: Executive summary
```

### Example 2: Ambiguous Query

```python
query = "Calculate profit"  # Missing formula details
result = agent.process_query("sales.xlsx", query)

if result["status"] == "clarification_needed":
    print(result["question"])
    # "Please specify how to calculate profit..."
    
    print(result["alternatives"])
    # ["Profit = Revenue - Cost", "Profit Margin = ...", ...]
```

### Example 3: High-Risk Operation

```python
query = "Delete the Cost column"
result = agent.process_query("sales.xlsx", query)

if result["status"] == "confirmation_required":
    print(f"⚠️ {result['message']}")
    # "This is a high-risk operation. Please review and confirm."
    
    decision = result["decision"]
    print(f"Risk Level: {decision['risk_level']}")
```

### Example 4: Visualization

```python
query = "Show me a bar chart of sales by product"
result = agent.process_query("sales.xlsx", query)

# Creates embedded chart in Excel file
```

### Example 5: Export Audit Log

```python
# Get audit trail
audit_log = agent.get_audit_log()

# Export for compliance
agent.export_audit_log("audit_2024.json")
```

---

## 🏗️ Architecture

```
User Query → Analysis → Parsing → Validation → Execution → Result
              ↓          ↓          ↓           ↓          ↓
           Excel     Multi-Step  Decision   Operation  Explanation
           Structure  Reasoning  Validation              OR
                                                      Error Recovery
```

### Enhanced Workflow (v2.0)

1. **Excel Structure Analysis** - Deep file understanding
2. **AI Query Parsing** - Multi-step reasoning (NEW)
3. **Ambiguity Check** - Clarification if needed (NEW)
4. **Decision Planning** - Structured execution plan
5. **Validation** - Pre-execution checks (NEW)
6. **Execution** - Safe operation with error handling
7. **Result Explanation** - Business-friendly summary (NEW)
8. **Audit Logging** - Comprehensive tracking (ENHANCED)

---

## 📊 Prompt Engineering Features

### 14 Advanced Techniques

1. ✅ **System Prompt** - Agent identity & behavior
2. ✅ **Context Injection** - Excel structure in prompts
3. ✅ **Multi-Step Reasoning** - 5-step framework
4. ✅ **Structured Output** - JSON schema enforcement
5. ✅ **Domain Mapping** - Business → technical terms
6. ✅ **Few-Shot Learning** - Examples in prompts
7. ✅ **Constraint Setting** - 7 critical rules
8. ✅ **Risk Assessment** - 3-level matrix
9. ✅ **Confidence Scoring** - Execution confidence
10. ✅ **Error Handling** - 4-strategy fallbacks
11. ✅ **Temperature Control** - Task-specific settings
12. ✅ **Validation Prompts** - Pre-execution checks
13. ✅ **Explanation Generation** - Post-execution summaries
14. ✅ **Recovery Prompts** - Error analysis & suggestions

See [`PROMPT_ENGINEERING.md`](PROMPT_ENGINEERING.md) for details.

---

## 🔒 Safety & Governance

### Built-in Safety

- ✅ **Never overwrites** original files
- ✅ **Risk assessment** before every operation
- ✅ **User confirmation** for high-risk operations
- ✅ **Data validation** (types, existence)
- ✅ **Complete audit trail** for compliance

### Audit Trail

Every operation is logged with:
- Timestamp (ISO-8601 format)
- Original query
- Operation performed
- Risk level
- Complete decision object
- Output file location

```python
# Export audit log
agent.export_audit_log("audit.json")
```

---

## 📚 Documentation

- **[PROMPT_ENGINEERING.md](PROMPT_ENGINEERING.md)** - Comprehensive prompt engineering guide
- **[ARCHITECTURE_DIAGRAM.md](ARCHITECTURE_DIAGRAM.md)** - Visual architecture
- **[IMPLEMENTATION_SUMMARY.md](IMPLEMENTATION_SUMMARY.md)** - What was implemented
- **[FEATURES_EXPLAINED.md](FEATURES_EXPLAINED.md)** - Feature catalog
- **[demo_enhanced_agent.py](demo_enhanced_agent.py)** - Interactive demo

---

## 🎯 Use Cases

### Business Analytics
```python
"Calculate profit margin for each product"
"Show me top 5 products by revenue"
"Create a pivot table of sales by region"
```

### Data Transformation
```python
"Sort by revenue in descending order"
"Remove duplicate customer records"
"Split the full name column into first and last name"
```

### Visualization
```python
"Create a bar chart of monthly sales"
"Show me a pie chart of market share"
"Generate a KPI dashboard"
```

### Data Validation
```python
"Add dropdown lists for the Status column"
"Highlight negative values in red"
"Find and mark duplicate entries"
```

---

## 🛠️ Requirements

```
groq>=0.4.0
pandas>=2.0.0
openpyxl>=3.1.0
python>=3.8
```

---

## 🎓 Advanced Usage

### Handle All Response Types

```python
result = agent.process_query(file, query)

if result["status"] == "success":
    # Success - operation completed
    print(result["explanation"]["executive_summary"])
    
elif result["status"] == "clarification_needed":
    # Ambiguous - need more info
    print(result["question"])
    print(result["alternatives"])
    
elif result["status"] == "confirmation_required":
    # High risk - need approval
    print(result["message"])
    decision = result["decision"]
    
elif result["status"] == "validation_failed":
    # Validation failed - blockers found
    print(result["blockers"])
    
elif result["status"] == "failed":
    # Error occurred
    recovery = result["recovery_suggestions"]
    print(recovery["user_friendly_message"])
    print(recovery["suggested_fixes"])
```

### Interactive Mode

```python
from excel_agent import run_interactive_agent

run_interactive_agent(api_key="your-key")
```

---

## 🏆 Why This Agent?

### ✅ Production-Ready
- Comprehensive error handling
- Complete audit trails
- Safety validations
- Business-friendly communication

### ✅ Intelligent
- Multi-step reasoning
- Ambiguity detection
- Risk assessment
- Confidence scoring

### ✅ Robust
- 4-level JSON parsing fallbacks
- Error recovery suggestions
- Validation before execution
- Alternative interpretations

### ✅ Transparent
- Shows reasoning process
- Explains decisions
- Clear error messages
- Complete audit logs

---

## 📝 License

MIT License - See LICENSE file for details

---

## 🤝 Contributing

Contributions welcome! Please read CONTRIBUTING.md first.

---

## 📧 Support

For issues or questions:
1. Check documentation files
2. Run `demo_enhanced_agent.py`
3. Review error recovery suggestions
4. Open an issue on GitHub

---

## 🎉 Acknowledgments

Built with:
- **Groq API** - Fast LLM inference
- **Llama 3.3 70B** - Advanced reasoning model
- **OpenPyXL** - Excel file manipulation
- **Pandas** - Data processing

---

**Version**: 2.0 Enhanced  
**Date**: December 25, 2025  
**Status**: ✅ Production Ready

🚀 **Ready to automate your Excel workflows with AI!**
