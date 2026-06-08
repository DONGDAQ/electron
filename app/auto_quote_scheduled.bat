@echo off
cd /d D:\baojia\electron\app
C:\Users\admin\AppData\Local\Python\pythoncore-3.14-64\python.exe auto_quote_scheduled.py >> auto_quote.log 2>&1
echo [%date% %time%] Done >> auto_quote.log
