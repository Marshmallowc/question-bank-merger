#!/bin/bash
cd "$(dirname "$0")"

echo "========================================"
echo "题库合并工具 - 新手版"
echo "========================================"
echo

# 检查Python是否安装
if ! command -v python3 &> /dev/null; then
    echo "错误：未检测到Python3！"
    echo "请先安装Python：https://www.python.org/downloads/"
    echo
    read -p "按回车键退出..."
    exit 1
fi

# 运行Python脚本
echo "正在启动题库合并工具..."
echo
python3 run.py