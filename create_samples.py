#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
创建示例Excel文件
"""
import pandas as pd
import os

def create_sample1():
    """创建第一个示例文件"""
    data = [
        ["这是一个示例Excel文件，展示典型的题库格式。", "", "", "", "", "", "", "", "", ""],
        ["题型", "题干", "正确答案", "解析", "分值", "难度系数", "选项A", "选项B", "选项C", "选项D"],
        ["单选题", "下列哪个是Python的特点？", "B", "Python是一种解释型、面向对象的高级编程语言", "1.0", "1", "编译型语言", "解释型语言", "汇编语言", "机器语言"],
        ["单选题", "Python中哪个关键字用于定义函数？", "def", "def是Python中定义函数的关键字", "1.0", "1", "function", "define", "func", "def"],
        ["多选题", "以下哪些是Python的数据类型？", "ABCD", "Python支持多种数据类型", "2.0", "2", "整数(int)", "字符串(str)", "列表(list)", "字典(dict)"],
        ["判断题", "Python是大小写敏感的。", "对", "Python中变量名是大小写敏感的", "1.0", "1"]
    ]

    df = pd.DataFrame(data)
    df.to_excel("examples/sample_questions/第一章示例_习题导出.xlsx", index=False, header=False)
    print("已创建：第一章示例_习题导出.xlsx")

def create_sample2():
    """创建第二个示例文件"""
    data = [
        ["这是一个示例Excel文件，展示典型的题库格式。", "", "", "", "", "", "", "", "", ""],
        ["题型", "题干", "正确答案", "解析", "分值", "难度系数", "选项A", "选项B", "选项C", "选项D"],
        ["单选题", "列表的索引从哪个数字开始？", "0", "Python列表索引从0开始", "1.0", "1", "0", "1", "-1", "2"],
        ["单选题", "如何获取列表的长度？", "C", "len()函数用于获取序列的长度", "1.0", "1", "length()", "size()", "len()", "count()"],
        ["判断题", "Python字典是有序的。", "错", "在Python 3.7之前，字典是无序的", "1.0", "1"]
    ]

    df = pd.DataFrame(data)
    df.to_excel("examples/sample_questions/第二章示例_习题导出.xlsx", index=False, header=False)
    print("已创建：第二章示例_习题导出.xlsx")

if __name__ == "__main__":
    # 确保目录存在
    os.makedirs("examples/sample_questions", exist_ok=True)

    create_sample1()
    create_sample2()
    print("\n示例文件创建完成！")