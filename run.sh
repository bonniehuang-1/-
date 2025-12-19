#!/bin/bash

echo "======================================================================"
echo "                  生产排产交付能力验证系统"
echo "======================================================================"
echo ""

# 检查Python是否安装
if ! command -v python3 &> /dev/null; then
    echo "[错误] 未检测到Python3，请先安装Python 3.8或更高版本"
    exit 1
fi

echo "[1/3] 检查依赖..."
if ! python3 -c "import pandas" &> /dev/null; then
    echo "[提示] 正在安装依赖包..."
    pip3 install -r requirements.txt
    if [ $? -ne 0 ]; then
        echo "[错误] 依赖安装失败"
        exit 1
    fi
fi

echo "[2/3] 检查数据文件..."
if [ ! -f "input/orders.xlsx" ]; then
    echo "[提示] 未找到数据文件，正在生成示例数据..."
    python3 create_sample_data.py
    if [ $? -ne 0 ]; then
        echo "[错误] 示例数据生成失败"
        exit 1
    fi
    echo ""
fi

echo "[3/3] 运行系统..."
echo ""
python3 main.py

echo ""
echo "======================================================================"
