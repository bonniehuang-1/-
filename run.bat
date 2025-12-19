@echo off
chcp 65001 >nul
echo ======================================================================
echo                  生产排产交付能力验证系统
echo ======================================================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到Python，请先安装Python 3.8或更高版本
    pause
    exit /b 1
)

echo [1/3] 检查依赖...
pip show pandas >nul 2>&1
if errorlevel 1 (
    echo [提示] 正在安装依赖包...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo [错误] 依赖安装失败
        pause
        exit /b 1
    )
)

echo [2/3] 检查数据文件...
if not exist "input\orders.xlsx" (
    echo [提示] 未找到数据文件，正在生成示例数据...
    python create_sample_data.py
    if errorlevel 1 (
        echo [错误] 示例数据生成失败
        pause
        exit /b 1
    )
    echo.
)

echo [3/3] 运行系统...
echo.
python main.py

echo.
echo ======================================================================
pause
