@echo off
chcp 65001 >nul
title 发票自动重命名和校验工具

echo ============================================
echo   发票自动重命名和校验工具
echo ============================================
echo.
echo  1. 重命名发票文件
echo  2. 校验发票信息
echo  3. 退出
echo.
set /p choice=请选择操作 (1/2/3):

if "%choice%"=="1" goto rename
if "%choice%"=="2" goto verify
if "%choice%"=="3" exit
echo 输入无效，请重新选择
pause
goto :eof

:rename
echo.
set /p pdf_dir=请输入PDF文件夹路径:
set /p excel_path=请输入合同号索引Excel路径:
echo.
echo 正在执行重命名...
uv run main.py rename --dir "%pdf_dir%" --excel "%excel_path%"
echo.
pause
goto :eof

:verify
echo.
set /p pdf_dir=请输入重命名后的PDF文件夹路径 (_Renamed):
set /p excel_path=请输入发票验证Excel路径:
echo.
echo 正在执行校验...
uv run main.py verify --dir "%pdf_dir%" --excel "%excel_path%"
echo.
pause
goto :eof
