$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
& "$scriptDir\.venv\Scripts\python.exe" -m streamlit run "$scriptDir\carvana_screen.py"
