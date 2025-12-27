x#!/bin/bash
# Ensures the script runs from the folder it is located in
cd -- "$(dirname "$0")"

# 1. Create a virtual environment if it doesn't exist
if [ ! -d "venv" ]; then
    echo "Setting up your environment for the first time..."
    python3 -m venv venv
    ./venv/bin/pip install streamlit canvasapi pandas openpyxl
fi

# 2. Launch the Streamlit app from the subfolder
echo "Launching Canvas Planner..."
./venv/bin/streamlit run app/appMAC.py --server.headless true