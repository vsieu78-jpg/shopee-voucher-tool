@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul 2>nul
cd /d "%~dp0"
title BUILD EXE PROTECTED - Ryan Nguyen v3.20

set "APP=ShopeeVoucherApiTester_ModernUI.py"
set "OUT=BUILD_OUTPUT"
set "RELEASE=RELEASE_TO_CUSTOMER"
set "PY_CMD="
set "PY_ARGS="
set "PY_VER="
set "COMPILER_FLAG="

if not exist "%APP%" (
    echo [LOI] Khong tim thay %APP%
    pause
    exit /b 1
)

if not exist "RYAN_VOUCHER.ico" (
    echo [LOI] Khong tim thay RYAN_VOUCHER.ico
    echo Hay dat file icon cung thu muc voi file BAT.
    pause
    exit /b 1
)

if not exist "mail_native_suite.py" (
    echo [LOI] Khong tim thay mail_native_suite.py
    pause
    exit /b 1
)
if not exist "outlook_native_backend.py" (
    echo [LOI] Khong tim thay outlook_native_backend.py
    pause
    exit /b 1
)

REM ============================================================
REM Python 3.13+ KHONG dung duoc --mingw64 voi Nuitka.
REM Uu tien 3.12/3.11 de co the dung MinGW64 tu dong.
REM Neu chi co 3.13 thi dung MSVC.
REM ============================================================
where py >nul 2>nul
if not errorlevel 1 (
    py -3.12 -c "import sys; print(sys.version)" >nul 2>nul && (
        set "PY_CMD=py"
        set "PY_ARGS=-3.12"
        set "PY_VER=3.12"
        set "COMPILER_FLAG=--mingw64"
        goto :py_ok
    )
    py -3.11 -c "import sys; print(sys.version)" >nul 2>nul && (
        set "PY_CMD=py"
        set "PY_ARGS=-3.11"
        set "PY_VER=3.11"
        set "COMPILER_FLAG=--mingw64"
        goto :py_ok
    )
    py -3.13 -c "import sys; print(sys.version)" >nul 2>nul && (
        set "PY_CMD=py"
        set "PY_ARGS=-3.13"
        set "PY_VER=3.13"
        set "COMPILER_FLAG=--msvc=latest"
        goto :py_ok
    )
)

where python >nul 2>nul
if not errorlevel 1 (
    for /f "tokens=*" %%V in ('python -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')" 2^>nul') do set "PY_VER=%%V"
    if "!PY_VER!"=="3.12" (
        set "PY_CMD=python"
        set "PY_ARGS="
        set "COMPILER_FLAG=--mingw64"
        goto :py_ok
    )
    if "!PY_VER!"=="3.11" (
        set "PY_CMD=python"
        set "PY_ARGS="
        set "COMPILER_FLAG=--mingw64"
        goto :py_ok
    )
    if "!PY_VER!"=="3.13" (
        set "PY_CMD=python"
        set "PY_ARGS="
        set "COMPILER_FLAG=--msvc=latest"
        goto :py_ok
    )
)

echo.
echo [LOI] Can Python 3.11, 3.12 hoac 3.13 64-bit.
pause
exit /b 1

:py_ok
echo ============================================================
echo   RYAN NGUYEN - NUITKA PROTECTED BUILD v3.20
echo ============================================================
echo [1/5] Python:
%PY_CMD% %PY_ARGS% --version
echo Compiler: %COMPILER_FLAG%

if "%PY_VER%"=="3.13" (
    echo.
    echo [THONG TIN] Python 3.13 bat buoc dung MSVC voi Nuitka.
    echo Neu may chua co Visual Studio C++ Build Tools, Nuitka se bao loi compiler.
    echo Khi do hay cai "Visual Studio 2022 Build Tools" voi workload "Desktop development with C++".
)

echo.
echo [2/5] Cai/cap nhat Nuitka va thu vien...
%PY_CMD% %PY_ARGS% -m pip install --upgrade pip
if errorlevel 1 goto :fail
%PY_CMD% %PY_ARGS% -m pip install --upgrade nuitka ordered-set zstandard "httpx[http2,socks]" openpyxl
if errorlevel 1 goto :fail

if exist "%OUT%" rmdir /s /q "%OUT%"
mkdir "%OUT%" >nul 2>nul

REM ============================================================
REM SECURITY SCAN: quet leak secret truoc khi build
REM ============================================================
echo.
echo [3a/5] Quet security leak truoc khi build...
findstr /r /i /m "VUBEL_API_KEY\|SECRET_SALT\|service_account\|admin_secret" "%APP%" >nul 2>nul
if not errorlevel 1 (
    echo.
    echo [BAO MAT] PHAT HIEN SECRET TRONG SOURCE CLIENT!
    echo   Tim thay: VUBEL_API_KEY hoac SECRET_SALT hoac service_account hoac admin_secret
    echo   HOAN TOAN HUY BUILD DE BAO VE KHACH HANG!
    echo.
    pause
    exit /b 1
)
findstr /r /i /m "VUBEL_API_KEY\|SECRET_SALT\|service_account\|admin_secret" "mail_native_suite.py" "outlook_native_backend.py" >nul 2>nul
if not errorlevel 1 (
    echo.
    echo [BAO MAT] PHAT HIEN SECRET TRONG MODULE BAO GOI!
    echo   HOAN TOAN HUY BUILD.
    echo.
    pause
    exit /b 1
)
echo [OK] Khong co secret bi ro ri. Tiep tuc build...

echo.
echo [3b/5] Dang bien dich native EXE bang Nuitka...
%PY_CMD% %PY_ARGS% -m nuitka ^
  --onefile ^
  --deployment ^
  --windows-console-mode=disable ^
  --enable-plugin=tk-inter ^
  --windows-icon-from-ico="RYAN_VOUCHER.ico" ^
  --include-data-files="RYAN_VOUCHER.ico=RYAN_VOUCHER.ico" ^
  --include-module=mail_native_suite ^
  --include-module=outlook_native_backend ^
  %COMPILER_FLAG% ^
  --assume-yes-for-downloads ^
  --lto=yes ^
  --python-flag=no_docstrings ^
  --python-flag=no_asserts ^
  --python-flag=isolated ^
  --include-package=httpx ^
  --include-package=httpcore ^
  --include-package=socksio ^
  --include-package=anyio ^
  --include-package=openpyxl ^
  --include-package-data=certifi ^
  --company-name="Ryan Nguyen" ^
  --product-name="Shopee Voucher Ryan Nguyen" ^
  --file-description="Shopee Voucher and SPC_ST Tool - Ryan Nguyen" ^
  --file-version=3.20.0.0 ^
  --product-version=3.20.0.0 ^
  --output-filename="ShopeeVoucher_RyanNguyen_v3.20.exe" ^
  --output-dir="%OUT%" ^
  "%APP%"
if errorlevel 1 goto :fail

echo.
echo [4/5] Quet .py/.pyc ro ri trong BUILD_OUTPUT...
set "LEAK_FOUND=0"
for /r "%OUT%" %%F in (*.py *.pyc) do (
    echo [CANH BAO] Phat hien source bi ro: %%F
    set "LEAK_FOUND=1"
)
if "!LEAK_FOUND!"=="1" (
    echo.
    echo [BAO MAT] Co file .py hoac .pyc trong BUILD_OUTPUT!
    echo Kiem tra lai cau hinh Nuitka --onefile.
    pause
    exit /b 1
)
echo [OK] Khong co .py/.pyc ro ri trong BUILD_OUTPUT.

echo.
echo [5/5] HOAN TAT.
if exist "%RELEASE%" rmdir /s /q "%RELEASE%"
mkdir "%RELEASE%" >nul 2>nul
copy /y "%OUT%\ShopeeVoucher_RyanNguyen_v3.20.exe" "%RELEASE%\ShopeeVoucher_RyanNguyen_v3.20.exe" >nul

echo.
echo EXE nam tai:
echo   %CD%\%OUT%\ShopeeVoucher_RyanNguyen_v3.20.exe
echo.
echo Thu muc de gui:
echo   %CD%\%RELEASE%
echo.
echo QUAN TRONG: chi gui EXE trong RELEASE_TO_CUSTOMER. KHONG GUI source.
explorer "%RELEASE%"
pause
exit /b 0

:fail
echo.
echo ============================================================
echo [LOI] Build that bai. Xem dong loi phia tren.
echo ============================================================
if "%PY_VER%"=="3.13" (
    echo Python 3.13 KHONG ho tro --mingw64 trong Nuitka.
    echo File nay da tu dong chuyen sang MSVC.
    echo Neu loi lien quan "cl.exe" / MSVC / Visual Studio:
    echo   Cai Visual Studio 2022 Build Tools
    echo   Chon workload: Desktop development with C++
    echo.
    echo Cach de nhat neu khong muon cai MSVC:
    echo   Cai Python 3.12 64-bit, sau do chay lai file BAT nay.
    echo   BAT se tu dong uu tien Python 3.12 + MinGW64.
)
pause
exit /b 1
