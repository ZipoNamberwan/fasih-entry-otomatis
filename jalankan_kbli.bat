@echo off
cd /d "%~dp0"

rem The script to run comes from this file's name: jalankan_kbli.bat -> kbli.py.
rem For a new script xxx.py, copy this file and rename the copy to jalankan_xxx.bat.
set "NAMA=%~n0"
set "NAMA=%NAMA:jalankan_=%"
title FASIH %NAMA%

if not exist "%~dp0%NAMA%.py" (
    echo Script %NAMA%.py tidak ditemukan di folder ini.
    pause
    exit /b 1
)

set "PY=%~dp0python\python.exe"
rem Python version used for python\ (official embeddable build from python.org)
set "PYTHON_VERSION=3.13.1"
rem Camoufox keeps its browser under LOCALAPPDATA; point it at browser\ in this folder instead
set "WIN_PD_OVERRIDE_LOCAL_APPDATA=%~dp0browser"
set "PYTHONUTF8=1"
rem Browser version the scripts were tested with (~500 MB download, ~1 GB on disk)
set "BROWSER_VERSION=official/152.0.4-beta.30"

rem 1) Python and the packages in requirements.txt: set up the first time only
call :python_siap
if errorlevel 1 (
    call :pasang_python
    call :python_siap
    if errorlevel 1 (
        echo.
        echo Gagal menyiapkan Python. Cek koneksi internet dan ruang disk, lalu jalankan lagi.
        rem Remove the half-finished folder so the next run starts clean
        rmdir /s /q "%~dp0python" 2>nul
        pause
        exit /b 1
    )
    echo.
)

rem 2) Camoufox browser: download the first time; afterwards the copy in browser\ is reused
call :browser_ada
if errorlevel 1 (
    echo Browser Camoufox belum ada. Mengunduh sekarang, hanya sekali.
    echo Butuh internet dan sekitar 1.5 GB ruang kosong di disk.
    echo.
    "%PY%" -m camoufox fetch %BROWSER_VERSION%
    rem "camoufox fetch" does not report failures in its exit code, so check again
    call :browser_ada
    if errorlevel 1 (
        echo.
        echo Gagal mengunduh browser Camoufox. Cek koneksi internet dan ruang disk,
        echo lalu jalankan lagi.
        pause
        exit /b 1
    )
    echo.
)

"%PY%" "%~dp0%NAMA%.py"

echo.
pause
exit /b


:python_siap
if not exist "%PY%" exit /b 1
"%PY%" -c "import camoufox, playwright, pandas, openpyxl" >nul 2>&1
exit /b


:pasang_python
echo Python belum siap. Mengunduh dan memasang sekarang, hanya sekali.
echo Butuh internet dan sekitar 500 MB ruang kosong di disk.
echo.
rmdir /s /q "%~dp0python" 2>nul
mkdir "%~dp0python"
echo [1/3] Mengunduh Python %PYTHON_VERSION%...
curl -fL# -o "%~dp0python\python.zip" "https://www.python.org/ftp/python/%PYTHON_VERSION%/python-%PYTHON_VERSION%-embed-amd64.zip" || exit /b 1
tar -xf "%~dp0python\python.zip" -C "%~dp0python" || exit /b 1
del "%~dp0python\python.zip"
rem Let the embeddable Python load packages from Lib\site-packages
for /f "delims=" %%f in ('dir /b "%~dp0python\python3*._pth"') do (
    > "%~dp0python\%%f" (
        echo python313.zip
        echo .
        echo Lib\site-packages
        echo import site
    )
)
echo [2/3] Memasang pip...
curl -fsSL -o "%~dp0python\get-pip.py" "https://bootstrap.pypa.io/get-pip.py" || exit /b 1
"%PY%" "%~dp0python\get-pip.py" --no-warn-script-location -q || exit /b 1
del "%~dp0python\get-pip.py"
echo [3/3] Memasang paket (camoufox, playwright, pandas, openpyxl)...
"%PY%" -m pip install --no-warn-script-location --disable-pip-version-check -q -r "%~dp0requirements.txt" || exit /b 1
exit /b 0


:browser_ada
"%PY%" -c "from camoufox.pkgman import camoufox_path; camoufox_path(download_if_missing=False)" >nul 2>&1
exit /b
