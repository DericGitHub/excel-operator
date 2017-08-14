@echo off
echo 'Installing python for Windows'
start /wait python-2.7.13.amd64.msi 
echo 'Installed python for Windows'
echo 'Installing PyQt for Windows'
start /wait PyQt4-4.10-gpl-Py2.7-Qt4.8.4-x64.exe
echo 'Installed PyQt for Windows'
echo 'Installing pywin32'
start /wait pywin32-220.win-amd64-py2.7.exe
echo 'Installed pywin32'
echo 'Installing openpyxl'
pip install openpyxl
echo 'Installed openpyxl'
echo 'Installing xlwings'
pip install xlwings
echo 'Installed xlwings for Windows'
pause
