@echo off
title UiPath to PAD Robin Pure LLM Converter
echo Starting Pure LLM Streamlit App...
cd /d "%~dp0"
python -m streamlit run app_llm.py --server.port 8503
pause

