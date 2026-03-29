@echo off
echo ==================================
echo   AI PPT Generator SaaS
echo ==================================
echo.
echo Installing dependencies...
pip install -r backend\requirements.txt
echo.
echo Starting server...
cd backend
python main.py
