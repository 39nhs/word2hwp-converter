@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "PY=python"
where python >nul 2>nul || set "PY=py -3"
%PY% --version >nul 2>nul
if errorlevel 1 (
  echo [!] 파이썬이 설치되어 있지 않습니다.
  echo     https://www.python.org 에서 파이썬을 설치한 뒤 다시 실행해 주세요.
  echo     설치할 때 "Add python.exe to PATH" 를 꼭 체크하세요.
  pause
  exit /b 1
)

%PY% -c "import hwpx, docx, lxml, tkinterdnd2, PIL" >nul 2>nul
if errorlevel 1 (
  echo 처음 실행이라 필요한 구성요소를 설치합니다. 잠시만 기다려 주세요...
  %PY% -m pip install --disable-pip-version-check -r requirements.txt
  if errorlevel 1 (
    echo [!] 설치에 실패했습니다. 인터넷 연결을 확인해 주세요.
    pause
    exit /b 1
  )
)

for /f "delims=" %%i in ('%PY% -c "import sys,os;print(os.path.join(os.path.dirname(sys.executable),'pythonw.exe'))"') do set "PYW=%%i"
if exist "%PYW%" (
  start "" "%PYW%" "%~dp0word2hwp_gui.pyw" %*
) else (
  %PY% "%~dp0word2hwp_gui.pyw" %*
)
