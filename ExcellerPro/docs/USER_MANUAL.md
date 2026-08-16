# Exceller Pro - User Manual

## Welcome

Exceller Pro is an AI-powered Excel assistant that helps you analyze data, create formulas, fix errors, and more. It uses Google Gemini AI to understand your questions and provide helpful answers.

---

## Quick Start

### Installation
1. Run `installer/install.bat` as Administrator
2. Open Excel
3. Press `Alt + F8`
4. Select `ExcellerPro_Settings` and click Run
5. Enter your API key

### Get Your API Key (Free)
1. Go to https://makersuite.google.com/app/apikey
2. Sign in with Google account
3. Click "Create API Key"
4. Copy the key

---

## Features

### AI Chat
Ask questions about your data in plain English.

**How to use:**
1. Select your data
2. Click "Ask AI" on toolbar
3. Type your question
4. Get instant answers

**Example questions:**
- "Summarize this sales data"
- "What trends do you see?"
- "Find any errors in this data"
- "Suggest improvements"

---

### Data Analysis
Get insights about your selected data.

**How to use:**
1. Select data range
2. Click "Analyze" on toolbar
3. Review the analysis

**You get:**
- Summary statistics
- Key insights
- Anomaly detection
- Recommendations

---

### Fix Formula
Repair broken Excel formulas.

**How to use:**
1. Click "Fix Formula" on toolbar
2. Enter the broken formula
3. Enter the error message
4. Get multiple fix options

**Example:**
- Formula: `=VLOOKUP(A1,Sheet2!A:B,2,FALSE)`
- Error: `#REF!`
- Get 3+ alternative solutions

---

### Price Calculator
Get help with price calculations.

**How to use:**
1. Click "Price Calc" on toolbar
2. Describe what you need
3. Get multiple calculation methods

**Examples:**
- "Calculate total with 10% discount"
- "Add 8% tax to subtotal"
- "Look up price from product list"

---

### Verify Formula
Check if a formula is correct.

**How to use:**
1. Select a cell with a formula
2. Click "Verify Formula" (Alt+F8)
3. Get verification report

**You get:**
- Syntax check
- Logic validation
- Edge case analysis
- Improvement suggestions

---

### Create Table
Format data as a professional table.

**How to use:**
1. Select raw data
2. Click "Create Table" (Alt+F8)
3. Get formatting suggestions

---

## Toolbar Reference

| Button | Command | Description |
|--------|---------|-------------|
| Ask AI | ExcellerPro_Chat | Chat with AI |
| Analyze | ExcellerPro_Analyze | Analyze data |
| Fix Formula | ExcellerPro_FixFormula | Fix broken formulas |
| Price Calc | ExcellerPro_PriceCalc | Price help |
| Settings | ExcellerPro_Settings | Configure app |

---

## All Commands (Alt+F8)

| Command | Description |
|---------|-------------|
| ExcellerPro_Chat | AI chat assistant |
| ExcellerPro_Analyze | Analyze selected data |
| ExcellerPro_FixFormula | Fix broken formulas |
| ExcellerPro_VerifyFormula | Verify formulas |
| ExcellerPro_CreateTable | Create formatted tables |
| ExcellerPro_PriceCalc | Price calculator |
| ExcellerPro_FormulaHelp | Formula suggestions |
| ExcellerPro_Settings | Change settings |
| ExcellerPro_Help | Show help |
| ExcellerPro_About | About info |
| ExcellerPro_Uninstall | Remove add-in |

---

## Troubleshooting

### "Macros disabled"
1. Go to File > Options > Trust Center
2. Click Trust Center Settings
3. Select Macro Settings
4. Choose "Enable all macros"
5. Check "Trust access to VBA project"

### "API Error"
1. Check internet connection
2. Verify API key is correct
3. Try again in a few minutes

### "Command not found"
1. Press Alt+F8
2. Look for Exceller commands
3. If missing, reinstall add-in

### Toolbar not showing
1. Go to View > Toolbars
2. Check "Exceller Pro"
3. Or run ExcellerPro_Chat to recreate it

---

## Support

- **Email:** support@exceller.app
- **Web:** https://exceller.app/help
- **GitHub:** https://github.com/exceller/excel-addin

---

## Privacy & Security

- Your data stays on your computer
- Only data you send to AI is processed
- API key stored locally in Windows
- No tracking or analytics
- No data sold to third parties

---

## License

Free to use for personal and commercial purposes.

---

## Changelog

### Version 3.0.0
- New: Price calculator
- New: Formula verification
- New: Improved error handling
- Fix: Better response formatting

### Version 2.0.0
- New: Custom toolbar
- New: Settings manager
- New: Help system
- Fix: API connection stability

### Version 1.0.0
- Initial release
- Basic AI chat
- Data analysis
- Formula help
