#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
题库合并工具 - 新手友好版
一键运行，自动检测和合并题库
"""
import os
import sys
import glob
import pandas as pd
from pathlib import Path

def detect_and_auto_merge():
    """自动检测并合并题库"""
    print("=" * 60)
    print("题库合并工具 - 新手友好版 v1.0")
    print("=" * 60)
    print("\n正在扫描当前目录的Excel文件...")

    # 查找所有Excel文件
    excel_files = glob.glob("*.xlsx") + glob.glob("*.xls")

    if not excel_files:
        print("\n❌ 未找到任何Excel文件！")
        print("请确保你的题库文件（.xlsx或.xls格式）放在当前目录下")
        input("\n按回车键退出...")
        return

    print(f"\n✅ 找到 {len(excel_files)} 个Excel文件：")
    for i, file in enumerate(excel_files, 1):
        print(f"  {i}. {file}")

    # 选择要合并的文件
    print("\n请选择要合并的文件：")
    print("1. 合并所有Excel文件")
    print("2. 只合并题库文件（包含'题'、'questions'、'章节'等关键词）")
    print("3. 手动选择文件")

    choice = input("\n请输入选项（1-3，默认为2）：").strip() or "2"

    # 筛选文件
    if choice == "1":
        selected_files = excel_files
        print("\n选择了所有Excel文件")
    elif choice == "2":
        keywords = ['题', 'questions', '章节', '章', 'chapter', 'quiz', 'test']
        selected_files = []
        for file in excel_files:
            if any(keyword in file.lower() for keyword in keywords):
                selected_files.append(file)
        if not selected_files:
            print("\n未找到题库文件，将使用所有Excel文件")
            selected_files = excel_files
        else:
            print(f"\n✅ 筛选出 {len(selected_files)} 个题库文件")
    else:
        print("\n请输入要合并的文件编号（用空格分隔）：")
        for i, file in enumerate(excel_files, 1):
            print(f"  {i}. {file}")
        numbers = input("\n文件编号：").strip().split()
        try:
            selected_files = [excel_files[int(n)-1] for n in numbers]
        except:
            print("\n输入错误，将使用所有文件")
            selected_files = excel_files

    print(f"\n即将合并 {len(selected_files)} 个文件...")

    # 分析第一个文件确定格式
    print("\n正在分析文件格式...")
    first_file = selected_files[0]
    format_type = detect_format(first_file)

    # 使用对应的配置合并文件
    success = merge_with_auto_config(selected_files, format_type)

    if success:
        print("\n✅ 合并成功！")
        print("\n生成的文件：")
        if os.path.exists("output/auto_merged.xlsx"):
            print("  📊 Excel文件: output/auto_merged.xlsx")
        if os.path.exists("output/auto_merged.docx"):
            print("  📝 Word文件: output/auto_merged.docx")

        # 询问是否打开文件
        open_file = input("\n是否打开生成的文件？(y/n): ").strip().lower()
        if open_file in ['y', 'yes', '是']:
            import subprocess
            import platform

            system = platform.system()
            if os.path.exists("output/auto_merged.xlsx"):
                if system == "Darwin":  # macOS
                    subprocess.run(["open", "output/auto_merged.xlsx"])
                elif system == "Windows":
                    os.startfile("output/auto_merged.xlsx")
                else:  # Linux
                    subprocess.run(["xdg-open", "output/auto_merged.xlsx"])
    else:
        print("\n❌ 合并失败，请检查文件格式")

    input("\n按回车键退出...")

def detect_format(filepath):
    """自动检测文件格式"""
    try:
        # 读取前3行来判断格式
        df = pd.read_excel(filepath, engine='openpyxl', header=None, nrows=3)

        # 检查第一行是否为说明文字
        first_row = df.iloc[0].astype(str).str.cat()
        if '为保证导出' in first_row or '格式' in first_row:
            return "chinese_style"  # 中文题库格式（第一行说明，第二行表头）
        else:
            # 检查是否有中文列名
            second_row = df.iloc[0].astype(str).str.cat()
            if '题型' in second_row or '题干' in second_row:
                return "chinese_direct"  # 中文格式（直接是表头）
            else:
                return "standard"  # 标准格式
    except:
        return "unknown"

def merge_with_auto_config(files, format_type):
    """使用自动配置合并文件"""
    try:
        # 导入合并器
        sys.path.insert(0, 'src')
        from merger import QuestionBankMerger

        # 根据格式选择配置
        if format_type == "chinese_style":
            config_file = "config/config.json"
        elif format_type == "chinese_direct":
            config_file = "config/config_standard.json"
        else:
            # 使用标准配置
            config_file = "config/config_standard.json"

        # 如果配置文件不存在，创建默认配置
        if not os.path.exists(config_file):
            create_default_config(config_file, format_type)

        # 创建合并器
        merger = QuestionBankMerger(config_file)
        merger.config["output_settings"]["excel_filename"] = "output/auto_merged.xlsx"
        merger.config["output_settings"]["word_filename"] = "output/auto_merged.docx"
        merger.config["file_patterns"] = files

        # 合并文件
        all_data = []
        for file in files:
            data = merger.read_excel_file(file)
            if not data.empty:
                all_data.append(data)

        if all_data:
            merged_data = pd.concat(all_data, ignore_index=True)
            merger.merged_data = merged_data

            # 保存文件
            merger.save_excel()
            if os.path.exists("output/auto_merged.xlsx"):
                print(f"✅ 成功合并 {len(merged_data)} 道题目")

                # 尝试保存Word文档
                try:
                    merger.save_word()
                except:
                    print("⚠️ Word文档生成失败，但Excel文件已成功生成")

                return True
        return False

    except Exception as e:
        print(f"❌ 错误: {e}")
        return False

def create_default_config(config_file, format_type):
    """创建默认配置文件"""
    import json

    if format_type == "chinese_style":
        config = {
            "excel_settings": {
                "has_header_row": True,
                "header_row_index": 1,
                "data_start_row": 2,
                "skip_description_row": True,
                "description_row_index": 0
            },
            "column_mapping": {
                "question_type": "题型",
                "question_text": "题干",
                "correct_answer": "正确答案",
                "analysis": "解析",
                "score": "分值",
                "difficulty": "难度系数",
                "options": ["选项A", "选项B", "选项C", "选项D", "选项E"]
            }
        }
    else:
        config = {
            "excel_settings": {
                "has_header_row": True,
                "header_row_index": 0,
                "data_start_row": 1,
                "skip_description_row": False
            },
            "column_mapping": {
                "question_type": "Question Type",
                "question_text": "Question",
                "correct_answer": "Answer",
                "analysis": "Analysis",
                "score": "Score",
                "difficulty": "Difficulty",
                "options": ["Option A", "Option B", "Option C", "Option D"]
            }
        }

    # 确保目录存在
    os.makedirs(os.path.dirname(config_file), exist_ok=True)

    with open(config_file, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)

def install_dependencies():
    """检查并安装依赖"""
    required = ['pandas', 'openpyxl']
    docx_required = 'python-docx'
    missing = []

    # 检查必需的依赖
    for module in required:
        try:
            __import__(module)
        except ImportError:
            missing.append(module)

    # 检查python-docx（用于生成Word）
    try:
        import docx
    except ImportError:
        missing.append(docx_required)

    if missing:
        print("\n需要安装以下依赖包：")
        for module in missing:
            print(f"  - {module}")

        # 在非交互式环境下自动安装
        if not sys.stdin.isatty():
            print("\n正在自动安装...")
        else:
            install = input("\n是否自动安装？(y/n): ").strip().lower()
            if install not in ['y', 'yes', '是']:
                print("\n请手动安装依赖：pip install pandas openpyxl python-docx")
                return False

        # 自动安装
        import subprocess
        try:
            subprocess.run([sys.executable, '-m', 'pip', 'install'] + missing, check=True)
            print("\n✅ 依赖安装完成！")
        except subprocess.CalledProcessError:
            print("\n❌ 自动安装失败，请手动安装")
            print(f"pip install {' '.join(missing)}")
            return False

    return True

if __name__ == "__main__":
    # 检查依赖
    if not install_dependencies():
        input("\n按回车键退出...")
        sys.exit(1)

    # 创建输出目录
    os.makedirs("output", exist_ok=True)

    # 运行主程序
    detect_and_auto_merge()