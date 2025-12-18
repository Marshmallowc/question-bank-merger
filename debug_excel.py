#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel文件调试工具
帮助用户分析Excel文件格式，生成对应的配置
"""
import pandas as pd
import sys
import os

def analyze_excel(filepath):
    """分析Excel文件格式"""
    print(f"\n分析文件: {filepath}")
    print("=" * 60)

    try:
        # 读取文件（不设置header，获取所有行）
        df = pd.read_excel(filepath, engine='openpyxl', header=None)

        print(f"\n文件基本信息:")
        print(f"  总行数: {len(df)}")
        print(f"  总列数: {len(df.columns)}")

        # 显示前5行的内容
        print("\n前5行内容:")
        for i in range(min(5, len(df))):
            print(f"\n第{i}行:")
            for j, cell in enumerate(df.iloc[i]):
                if pd.notna(cell):
                    # 截断过长的内容
                    content = str(cell)
                    if len(content) > 50:
                        content = content[:50] + "..."
                    print(f"  列{j}: {content}")
                else:
                    print(f"  列{j}: (空)")

        # 智能检测表头位置
        print("\n智能检测:")
        header_candidates = []

        for i in range(min(5, len(df))):
            row = df.iloc[i]
            # 检查是否包含典型的列名关键词
            keywords = ['题型', '题干', '问题', 'question', 'answer', '答案', '选项', 'option']
            score = 0
            for cell in row:
                if pd.notna(cell):
                    cell_str = str(cell).lower()
                    for keyword in keywords:
                        if keyword in cell_str:
                            score += 1
            if score > 0:
                header_candidates.append((i, score))

        if header_candidates:
            header_candidates.sort(key=lambda x: x[1], reverse=True)
            print(f"  可能的表头位置: 第{header_candidates[0][0]}行 (匹配度: {header_candidates[0][1]}个关键词)")

            # 推荐配置
            header_row = header_candidates[0][0]
            print(f"\n推荐配置:")
            print(f"  \"header_row_index\": {header_row},")
            print(f"  \"data_start_row\": {header_row + 1},")
            print(f"  \"skip_description_row\": {header_row > 0},")

            # 显示检测到的列名
            print(f"\n检测到的列名:")
            for i, cell in enumerate(df.iloc[header_row]):
                if pd.notna(cell):
                    print(f"  列{i}: {cell}")
        else:
            print("  未检测到明显的表头行")
            print("  请手动检查文件内容")

        # 生成配置建议
        print("\n配置建议:")
        print("1. 如果第一行是说明文字，设置:")
        print("   \"skip_description_row\": true,")
        print("   \"description_row_index\": 0,")
        print("   \"header_row_index\": 1,")
        print("   \"data_start_row\": 2")
        print("\n2. 如果第一行就是表头，设置:")
        print("   \"skip_description_row\": false,")
        print("   \"header_row_index\": 0,")
        print("   \"data_start_row\": 1")

        return True

    except Exception as e:
        print(f"\n错误: 无法读取文件")
        print(f"错误信息: {e}")
        return False

def generate_config_template():
    """生成配置文件模板"""
    template = {
        "excel_settings": {
            "has_header_row": True,
            "header_row_index": 1,  # 根据实际情况修改
            "data_start_row": 2,    # 根据实际情况修改
            "skip_description_row": True,  # 根据实际情况修改
            "description_row_index": 0
        },
        "column_mapping": {
            "question_type": "题型",  # 修改为实际的列名
            "question_text": "题干",  # 修改为实际的列名
            "correct_answer": "正确答案",  # 修改为实际的列名
            "analysis": "解析",
            "score": "分值",
            "difficulty": "难度系数",
            "options": ["选项A", "选项B", "选项C", "选项D", "选项E"]  # 修改为实际的列名
        },
        "output_settings": {
            "excel_filename": "output/merged_questions.xlsx",
            "word_filename": "output/merged_questions.docx",
            "include_analysis": True,
            "include_difficulty": True
        },
        "file_patterns": [
            "*.xlsx"
        ]
    }

    import json
    with open("config_template.json", "w", encoding="utf-8") as f:
        json.dump(template, f, ensure_ascii=False, indent=2)

    print("\n已生成配置模板: config_template.json")
    print("请根据实际Excel文件格式修改此文件")

def main():
    """主函数"""
    if len(sys.argv) < 2:
        print("使用方法: python debug_excel.py <excel文件路径>")
        print("示例: python debug_excel.py 我的题库.xlsx")
        sys.exit(1)

    filepath = sys.argv[1]

    if not os.path.exists(filepath):
        print(f"错误: 文件 '{filepath}' 不存在")
        sys.exit(1)

    # 分析文件
    success = analyze_excel(filepath)

    if success:
        # 询问是否生成配置模板
        print("\n" + "=" * 60)
        response = input("\n是否生成配置文件模板？(y/n): ").strip().lower()
        if response in ['y', 'yes', '是']:
            generate_config_template()

if __name__ == "__main__":
    main()