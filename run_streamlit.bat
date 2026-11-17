@echo off
title NUPKG Streamlit Browser
echo Starting Streamlit NUPKG Browser...
cd /d "%~dp0"
python -m streamlit run nupkg_browser.py --server.port 8502
pause
