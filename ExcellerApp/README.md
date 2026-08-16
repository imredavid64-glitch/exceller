# Exceller - AI Excel Assistant

A Grammarly-like desktop assistant for Microsoft Excel, powered by Google Gemini AI.

![Exceller](https://img.shields.io/badge/Exceller-AI%20Excel%20Assistant-4CAF50?style=for-the-badge)
![Python](https://img.shields.io/badge/Python-3.8+-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Gemini](https://img.shields.io/badge/Gemini-AI-4285F4?style=for-the-badge&logo=google&logoColor=white)

## Features

- **Floating Assistant** - Always-on-top window that works alongside Excel
- **AI Chat** - Natural language queries about your data
- **File Monitoring** - Auto-detects changes to Excel files
- **Quick Actions** - One-click analysis, statistics, and data cleaning
- **Dark Theme** - Modern, eye-friendly interface
- **System Tray** - Runs in background, accessible anytime

## Quick Start

### Windows

1. Run `install.bat` as administrator
2. Wait for installation to complete
3. Run `run.bat` to start Exceller

### Manual Installation

```bash
# Create virtual environment
python -m venv venv

# Activate (Windows)
venv\Scripts\activate

# Activate (Mac/Linux)
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Run the app
python main.py
```

## How to Use

### 1. Start Exceller
- Double-click `run.bat` or run `python main.py`
- The app appears in your system tray

### 2. Load an Excel File
- Click the floating assistant window
- Click "Browse" to select an Excel file
- Or drag and drop a file onto the window

### 3. Ask Questions
- Type questions in the input field
- Use quick action buttons for common tasks
- The AI analyzes your data and provides insights

### 4. Quick Actions

| Button | Action |
|--------|--------|
| 📈 Analyze | Get insights about your data |
| 🔍 Find Issues | Detect errors and inconsistencies |
| 📊 Statistics | Calculate key statistics |
| 🧹 Clean | Get data cleaning suggestions |

## Example Prompts

```
"Summarize this sales data"
"Find duplicates in column A"
"Write a formula to calculate totals"
"Create a pivot table by region"
"Clean this messy data"
"Generate a chart for this data"
"Explain what this data shows"
```

## Features in Detail

### AI-Powered Analysis
- Natural language understanding
- Context-aware responses
- Formula generation
- VBA code creation

### File Monitoring
- Watches for file changes
- Auto-refreshes data
- Supports .xlsx, .xls, .csv

### Data Processing
- Reads Excel and CSV files
- Extracts statistics
- Identifies patterns and issues

## System Requirements

- Windows 10/11, macOS, or Linux
- Python 3.8 or higher
- 4GB RAM recommended
- Internet connection (for AI features)

## Privacy & Security

- Your data stays on your computer
- Only sends data to Google Gemini API when you ask questions
- API key stored locally
- No data collection or tracking

## API Key

The app uses the Google Gemini API. Set your key with the
`GEMINI_API_KEY` environment variable (or pass it directly to
`AIEngine(api_key="...")`). No key is embedded in the source code.

To use your own key:
1. Get a key at https://makersuite.google.com/app/apikey
2. Windows: `set GEMINI_API_KEY=your-key` (or add it in System Properties)
   Mac/Linux: `export GEMINI_API_KEY=your-key`
3. Run the app

## Troubleshooting

### "Python not found"
- Install Python from https://python.org
- Check "Add Python to PATH" during installation

### "Module not found"
- Run `install.bat` again
- Or run `pip install -r requirements.txt`

### "API Error"
- Check your internet connection
- Verify API key is valid

### Window not showing
- Check system tray for Exceller icon
- Double-click the tray icon to show

## Development

### Project Structure

```
ExcellerApp/
├── main.py              # Main application
├── modules/
│   ├── __init__.py
│   ├── ai_engine.py     # Gemini AI integration
│   ├── excel_parser.py  # Excel file handling
│   ├── file_monitor.py  # File change detection
│   └── theme.py         # UI theming
├── assets/              # Icons and resources
├── config/              # Configuration files
├── requirements.txt     # Python dependencies
├── install.bat          # Windows installer
└── run.bat              # Windows launcher
```

### Adding Features

1. Edit `main.py` for UI changes
2. Edit `modules/ai_engine.py` for AI features
3. Edit `modules/excel_parser.py` for file handling

## License

MIT License - Free to use and modify

## Support

For issues or questions:
- Check the troubleshooting section
- Review the code comments
- Test with sample Excel files
