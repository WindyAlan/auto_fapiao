@echo off
cd /d "%~dp0"
title Invoice Auto-Rename and Verify Tool

echo ============================================
echo   Invoice Auto-Rename and Verify Tool
echo ============================================
echo.
echo  1. Rename invoice PDFs
echo  2. Verify invoice data
echo  3. Exit
echo.
set /p choice="Select (1/2/3): "

if "%choice%"=="1" goto rename
if "%choice%"=="2" goto verify
if "%choice%"=="3" goto end
echo Invalid choice.
pause
goto end

:rename
echo.
set /p pdf_dir="PDF folder path: "
set /p excel_path="Contract index Excel path: "
echo.
echo Running rename...
uv run main.py rename --dir "%pdf_dir%" --excel "%excel_path%"
echo.
pause
goto end

:verify
echo.
set /p pdf_dir="Renamed PDF folder path (_Renamed): "
set /p excel_path="Invoice verify Excel path: "
echo.
echo Running verify...
uv run main.py verify --dir "%pdf_dir%" --excel "%excel_path%"
echo.
pause
goto end

:end
