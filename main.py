#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
题库合并工具 - 主入口
简单的命令行界面
"""
import os
import sys

# 添加src目录到Python路径
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from merger import QuestionBankMerger

def main():
    """简单的交互式界面"""
    print("=" * 50)
    print("题库合并工具 v1.0")
    print("=" * 50)
    print()

    # 选择配置
    print("请选择配置：")
    print("1. 中文题库格式（默认）")
    print("2. 标准Excel格式")
    print("3. 自定义配置")

    choice = input("请输入选项 (1-3): ").strip()

    if choice == "2":
        config_path = "config/config_standard.json"
    elif choice == "3":
        config_path = input("请输入配置文件路径: ").strip()
    else:
        config_path = "config/config.json"

    # 输入目录
    input_dir = input("请输入题库文件目录（默认为当前目录）: ").strip()
    if not input_dir:
        input_dir = "."

    # 创建合并器
    merger = QuestionBankMerger(config_path)

    # 合并文件
    print("\n开始合并题库...")
    data = merger.merge_files(input_dir)

    if data.empty:
        print("\n未找到任何题库文件！")
        return

    # 生成报告
    report = merger.generate_report()
    print("\n" + "=" * 50)
    print("合并报告")
    print("=" * 50)

    for key, value in report.items():
        print(f"\n{key}:")
        if isinstance(value, dict):
            for k, v in value.items():
                print(f"  {k}: {v}")
        else:
            print(f"  {value}")

    # 保存文件
    print("\n正在保存文件...")
    merger.save_excel()
    merger.save_word()

    print("\n[SUCCESS] 完成！文件已保存到 output/ 目录")

if __name__ == "__main__":
    main()