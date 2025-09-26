chcp 65001
@echo off
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed. Installing Python...
    REM 下載 Python 安裝檔
    powershell -Command "Invoke-WebRequest -Uri https://www.python.org/ftp/python/3.12.4/python-3.12.4-amd64.exe -OutFile python-installer.exe"
    REM 執行安裝
    start /wait python-installer.exe /quiet InstallAllUsers=1 PrependPath=1
    REM 刪除安裝檔
    del python-installer.exe
) else (
    echo 已安裝Python.
)

echo 檢查更新 pip
python -m pip install --upgrade pip

REM --- 安裝 sv_ttk ---
pip list | findstr sv_ttk >nul 2>&1
if %errorlevel% neq 0 (
    echo 安裝 sv_ttk...
    pip install sv_ttk
) else (
    echo sv_ttk 已安裝.
)

REM --- 安裝 darkdetect ---
pip list | findstr darkdetect >nul 2>&1
if %errorlevel% neq 0 (
    echo 安裝 darkdetect...
    pip install darkdetect
) else (
    echo darkdetect 已安裝.
)

REM --- 安裝 pandas ---
pip list | findstr pandas >nul 2>&1
if %errorlevel% neq 0 (
    echo 安裝 pandas...
    pip install pandas
) else (
    echo pandas 已安裝.
)

REM --- 安裝 selenium ---
pip list | findstr selenium >nul 2>&1
if %errorlevel% neq 0 (
    echo 安裝 selenium...
    pip install selenium
) else (
    echo selenium 已安裝.
)

REM --- 安裝 pywinstyles ---
pip list | findstr pywinstyles >nul 2>&1
if %errorlevel% neq 0 (
    echo 安裝 pywinstyles...
    pip install pywinstyles
) else (
    echo pywinstyles 已安裝.
)

REM --- 安裝 openpyxl ---
pip list | findstr openpyxl >nul 2>&1
if %errorlevel% neq 0 (
    echo 安裝 openpyxl...
    pip install openpyxl
) else (
    echo openpyxl 已安裝.
)

echo 完成套件檢查.

pause

