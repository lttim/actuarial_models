@echo off
cd /d "%~dp0"
if not exist reports mkdir reports
python -m pytest --html=reports\pytest_report.html --self-contained-html
if exist reports\pytest_report.html start "" "%~dp0reports\pytest_report.html"
exit /b %ERRORLEVEL%
