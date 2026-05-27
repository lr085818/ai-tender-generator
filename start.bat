@echo off
chcp 65001 >nul
title AI智能招投标文档生成系统

echo ============================================
echo    AI智能招投标文档生成系统
echo    启动中...
echo ============================================
echo.

:: 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到Python，请先安装Python 3.9+
    pause
    exit /b 1
)

:: 检查.env文件
if not exist ".env" (
    echo [错误] 未找到.env配置文件，请先配置API Key
    pause
    exit /b 1
)

:: 创建必要目录
if not exist "outputs" mkdir outputs
if not exist "uploads" mkdir uploads

:: 安装依赖（首次运行）
if not exist ".deps_installed" (
    echo [信息] 首次运行，正在安装依赖...
    pip install flask flask-cors python-docx mammoth PyPDF2 python-dotenv requests -q
    echo. > .deps_installed
    echo [信息] 依赖安装完成
)

echo.
echo [信息] 服务启动中...
echo [信息] 访问地址: http://localhost:3012
echo [信息] 按 Ctrl+C 停止服务
echo.

python main.py
pause
