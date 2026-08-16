# Excel Assistant Setup Guide

## Quick Start (Windows, ~30 seconds)

### Step 1: Close Excel
Close all Microsoft Excel windows (version 2016 or later)

### Step 2: Run the installer
Double-click `install.bat` (or run `install.ps1` from PowerShell).
The installer automatically:
1. Compiles `ExcelAssistant.xlam` from `ExcelAssistant.bas` and `frmAssistant.frm`
   (it drives Excel itself via COM automation - you never open the VBA editor)
2. Copies the add-in to `%APPDATA%\Microsoft\AddIns`
3. Registers it so Excel loads it automatically on every start

### Step 3: Start Using
Excel opens automatically. The **Excel Assistant** toolbar appears
at the top of Excel - no manual setup or macro running needed.

---

## Features Overview

### AI Chat
Click the **AI Chat** button to open a conversation with the AI assistant.
- Ask questions about your data
- Get help with formulas
- Request data analysis

### Analyze Selection
1. Select a range of data in Excel
2. Click **Analyze**
3. Get insights and statistics

### Create Table
1. Select raw data
2. Click **Create Table**
3. Get a formatted, professional table

### Check Files
1. Click **Check Files**
2. Select a folder
3. Get analysis of files in that folder

---

## Usage Examples

### Example 1: Analyze Sales Data
1. Select your sales data (including headers)
2. Click **Analyze**
3. The AI will provide:
   - Summary statistics
   - Key insights
   - Recommendations

### Example 2: Get Formula Help
1. Click **AI Chat**
2. Type: "Write a formula to calculate total sales for January"
3. Get the formula with explanation

### Example 3: Clean Data
1. Select messy data
2. Click **AI Chat**
3. Type: "Help me clean this data"
4. Get step-by-step instructions

### Example 4: Create Pivot Table
1. Click **AI Chat**
2. Type: "Create a pivot table showing sales by region"
3. Get the VBA code or instructions

---

## Keyboard Shortcuts

| Action | Shortcut |
|--------|----------|
| Open AI Chat | Click toolbar button |
| Analyze Selection | Click toolbar button |
| Create Table | Click toolbar button |
| Check Files | Click toolbar button |
| Settings | Click toolbar button |

---

## Troubleshooting

### "Can't find project or library"
1. Open VBA Editor (`Alt + F11`)
2. Go to **Tools > References**
3. Check these boxes:
   - Microsoft Excel Object Library
   - Microsoft Visual Basic for Applications Extensibility
4. Click OK

### "Permission denied" Error
1. Go to **File > Options > Trust Center**
2. Click **Trust Center Settings**
3. Go to **Macro Settings**
4. Select "Enable all macros"
5. Check "Trust access to the VBA project object model"

### API Key Not Working
1. Click **Settings** in the toolbar
2. Enter your Gemini API key
3. Get your key at: https://makersuite.google.com/app/apikey

### Toolbar Not Showing
1. Right-click the toolbar area
2. Select "Excel Assistant"
3. Or run `SetupAssistant` again

---

## Advanced Features

### Custom Formulas
You can use the AI to generate custom formulas:
1. Click **AI Chat**
2. Describe what you need
3. Copy the formula to your cell

### Batch Processing
Process multiple sheets:
1. Click **AI Chat**
2. Type: "Analyze all sheets in this workbook"
3. Get a comprehensive report

### Data Export
1. Select data
2. Click **AI Chat**
3. Type: "Export this data to CSV"
4. Save the generated file

---

## Getting Help

### In-App Help
- Click **AI Chat** and ask any question
- The AI can help with Excel features

### Report Issues
If you encounter issues:
1. Check the troubleshooting section
2. Verify your API key is valid
3. Ensure you have internet connection

---

## API Key Security

Your Gemini API key is:
- Stored locally on your computer
- Never sent to third parties
- Used only for AI requests

To update your key:
1. Click **Settings**
2. Enter new key
3. Click OK

---

## Requirements

- Microsoft Excel 2016 or later (Windows)
- Internet connection
- Google Gemini API key (free tier available)

---

## File Locations

After installation, the add-in is stored in:
- Windows: `%APPDATA%\Microsoft\AddIns\ExcelAssistant.xlam`
- The VBA code is compiled inside that add-in file
- Settings are stored in Windows Registry

## Uninstall

To remove Excel Assistant:
1. Open Excel
2. Go to **File > Options > Add-ins**
3. Uncheck **Excel Assistant**
4. Click OK
5. Delete `ExcelAssistant.xlam` from the AddIns folder

