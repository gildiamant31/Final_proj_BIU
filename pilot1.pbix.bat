::can using this command for exacute powershell script which install python3.6 on computer / VM (Windows)
:: checking if python3 is installed on it& install it if it required. 
python --version 3>NUL
if errorlevel 1 Powershell.exe -File src\Install-Python.ps1

:: install all relavent packages.
pip install openpyxl xlsxwriter xlrd tk biopython
pip install pandas
::run the python script for editing the excel.
python src\updateExcel.py 2>> python.log 
:: open Power BI desktop with the specific file.
src\pilot2.pbix
ECHO "completed"