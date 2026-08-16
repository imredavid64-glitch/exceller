# Excel Assistant - Gemini AI Powered

A Microsoft Excel add-in powered by Google's Gemini AI that helps you analyze data, create tables, check files, and more.

## Features

- **AI Chat**: Ask questions about your data in natural language
- **Data Analysis**: Get insights, summaries, and statistics from selected data
- **Table Creation**: Generate formatted tables from raw data
- **File Analysis**: Analyze CSV, text, and other files
- **Formula Generation**: Get help writing Excel formulas
- **Data Cleaning**: Identify and fix data quality issues
- **Chart Suggestions**: Get recommendations for data visualization

## Installation

### Quick Install (Windows)

1. Close all Excel windows
2. Open the folder where this file is located
3. Double-click `install.bat`
4. Open Excel and look for "Excel Assistant" in the ribbon

### Manual Install (All Platforms)

1. Open Excel
2. Press `Alt + F11` to open VBA Editor
3. Go to `File > Import File`
4. Select `ExcelAssistant.bas`
5. Close VBA Editor
6. Press `Alt + F8`, select `SetupAssistant`, click `Run`

## Usage

1. Select data in your spreadsheet
2. Click "Excel Assistant" in the ribbon
3. Use the task pane to interact with the AI
4. Ask questions like:
   - "Summarize this data"
   - "Create a pivot table"
   - "Find duplicates"
   - "Generate a chart"
   - "Clean this data"
   - "Write a formula to calculate..."

## API Key

The add-in uses Google's Gemini API. Your API key is stored locally and never shared.

## Requirements

- Microsoft Excel 2016 or later (Windows/Mac)
- Internet connection
- Google Gemini API key

## Privacy

- Your data stays on your computer
- Only the data you explicitly send to the AI is processed
- No data is stored on external servers
- API key is stored locally in Excel's trusted settings
