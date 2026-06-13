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
pip3 install flask python-docx openpyxl requests --quiet 2>/dev/null || \
pip install flask python-docx openpyxl requests --quiet 2>/dev/null

echo ""
echo "Which app do you want to start?"
echo "  1) Web Form (manual entry)    → http://localhost:5000"
echo "  2) Legal Dashboard (HubSpot)  → http://localhost:5001"
echo ""
read -p "Enter 1 or 2 (default: 1): " choice

if [ "$choice" = "2" ]; then
    # Check HubSpot token
    if [ -f ".env" ] && grep -q "HUBSPOT_TOKEN=your_hubspot" .env 2>/dev/null; then
        echo ""
        echo "⚠  WARNING: HUBSPOT_TOKEN is not set in your .env file."
        echo "   Please edit .env and paste your HubSpot Private App Token."
        echo ""
        read -p "Press Enter to open .env for editing, or Ctrl+C to cancel..."
        open -e .env
        echo "After saving .env, re-run this launcher."
        exit 0
    fi
    echo ""
    echo "Starting Legal Dashboard (HubSpot)..."
    echo "When you see 'Running on http://127.0.0.1:5001',"
    echo "open your browser and go to: http://localhost:5001"
    echo ""
    echo "(To stop the app, close this window)"
    echo ""
    sleep 3 && open "http://localhost:5001" &
    python3 hubspot_app.py
else
    echo ""
    echo "Starting Web Form..."
    echo "When you see 'Running on http://127.0.0.1:5000',"
    echo "open your browser and go to: http://localhost:5000"
    echo ""
    echo "(To stop the app, close this window)"
    echo ""
    sleep 3 && open "http://localhost:5000" &
    python3 web_app.py
fi
