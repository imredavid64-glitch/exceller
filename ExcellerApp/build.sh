#!/bin/bash
echo "========================================"
echo "Building Exceller - Standalone App"
echo "========================================"
echo ""

# Check Python
if ! command -v python3 &> /dev/null; then
    echo "Python3 not found! Install from python.org"
    exit 1
fi

# Install dependencies
echo "Installing dependencies..."
pip3 install -r requirements.txt
pip3 install pyinstaller

echo ""
echo "Building Exceller..."
echo "This may take 2-3 minutes..."
echo ""

pyinstaller --onefile --windowed --name "Exceller" main.py

echo ""
echo "========================================"
echo "Build Complete!"
echo "========================================"
echo ""
echo "Find Exceller in the 'dist' folder"
echo ""
