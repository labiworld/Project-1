#!/bin/bash
# LandCulture Contract Generator — One-click launcher for Mac
# Double-click this file to start the app

# Go to the folder where this script lives
cd "$(dirname "$0")"

echo "======================================"
echo " LandCulture Contract Generator"
echo "======================================"
echo ""

# Install dependencies if not already installed
echo "Checking dependencies..."
pip3 install flask python-docx openpyxl --quiet 2>/dev/null || \
pip install flask python-docx openpyxl --quiet 2>/dev/null

echo ""
echo "Starting the app..."
echo "When you see 'Running on http://127.0.0.1:5000',"
echo "open your browser and go to: http://localhost:5000"
echo ""
echo "(To stop the app, close this window)"
echo ""

# Open browser automatically after 3 seconds
sleep 3 && open "http://localhost:5000" &

# Start the web app
python3 web_app.py
