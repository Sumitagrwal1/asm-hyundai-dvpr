@echo off
python "%~dp0generate_dashboard.py"
start "" "%~dp0DVPR_Live_Dashboard.html"
