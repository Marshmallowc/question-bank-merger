# 题库合并工具

一个强大的题库合并工具，可以将多个Excel格式的题库文件合并成统一的文档。专为教育工作者、培训师和需要整合分散题库的教育机构设计。

## 概述

题库合并工具能够自动检测并合并包含题目数据的Excel文件，支持多种格式，可生成Word和Excel两种输出格式。无需任何编程知识即可使用。

## 功能特性

- **智能检测**：自动识别Excel文件格式和列结构
- **一键操作**：运行工具即可自动完成所有工作
- **多种输出格式**：生成Excel和Word文档
- **灵活配置**：通过JSON配置文件支持各种题库格式
- **错误处理**：完善的错误报告和故障排除指导
- **中文支持**：针对中文教育题库格式进行优化

## 快速开始

### 环境要求

- Python 3.7 或更高版本
- 必需的包：pandas, openpyxl, python-docx

### 安装

1. 克隆仓库
   ```bash
   git clone https://github.com/Marshmallowc/question-bank-merger.git
   cd question-bank-merger
   ```

2. 安装依赖
   ```bash
   pip install pandas openpyxl python-docx
   ```

3. 将Excel题库文件放入项目目录

4. 运行合并工具
   ```bash
   python run.py
   ```

## 使用方法

### 基础使用（推荐新手使用）

工具提供交互式界面，引导您完成整个过程：

1. 将Excel文件放入项目目录
2. 运行 `python run.py`
3. 按照屏幕提示选择要合并的文件
4. 生成的文件将保存在 `output/` 目录中

### 高级使用

如需更多控制，可以使用命令行参数：

```bash
# 使用默认配置
python src/merger.py

# 指定自定义配置
python src/merger.py --config config/my_config.json

# 指定输入目录和文件模式
python src/merger.py --input /path/to/questions --pattern "chapter*.xlsx"

# 仅生成Word文档
python src/merger.py --word-only

# 仅生成Excel文档
python src/merger.py --excel-only
```

## 配置说明

工具通过JSON配置文件支持多种Excel格式。包含的默认配置：

- `config/config.json` - 中文题库格式（包含表头说明行）
- `config/config_standard.json` - 标准Excel格式

### 自定义配置

创建自定义JSON配置文件以匹配特定格式：

```json
{
  "excel_settings": {
    "has_header_row": true,
    "header_row_index": 1,
    "data_start_row": 2,
    "skip_description_row": true
  },
  "column_mapping": {
    "question_type": "题型",
    "question_text": "题干",
    "correct_answer": "正确答案",
    "analysis": "解析",
    "score": "分值",
    "difficulty": "难度系数",
    "options": ["选项A", "选项B", "选项C", "选项D", "选项E"]
  },
  "output_settings": {
    "excel_filename": "output/merged_questions.xlsx",
    "word_filename": "output/merged_questions.docx"
  }
}
```

## 支持的题型

- 单选题
- 判断题
- 填空题
- 简答题
- 论述题
- 自定义题型

## 故障排除

### 常见问题

**文件未检测到**
- 确保Excel文件在正确的目录中
- 检查文件扩展名（.xlsx 或 .xls）
- 确认文件没有密码保护

**合并失败**
- 尝试将Excel文件另存为.xlsx格式
- 检查文件是否损坏
- 使用调试工具：`python debug_excel.py your_file.xlsx`

**依赖安装错误**
- 更新pip：`python -m pip install --upgrade pip`
- 如需要可使用备用源：
  ```bash
  pip install -i https://pypi.tuna.tsinghua.edu.cn/simple/ pandas openpyxl python-docx
  ```

## 使用示例

### 示例1：大学题库

您的Excel文件可能如下所示：

| 题型 | 题干 | 正确答案 | 选项A | 选项B | 选项C | 选项D |
|------|------|----------|-------|-------|-------|-------|
| 单选题 | 下列哪个是Python的特点？ | B | 编译型语言 | 解释型语言 | 汇编语言 | 机器语言 |

工具将自动检测此格式并合并多个类似文件。

### 示例2：英文题目

| Question Type | Question | Answer | Option A | Option B | Option C | Option D |
|---------------|----------|--------|----------|----------|----------|----------|
| Multiple Choice | What is 2+2? | B | 3 | 4 | 5 | 6 |

## 项目结构

```
question-bank-merger/
├── src/
│   └── merger.py          # 核心合并逻辑
├── config/
│   ├── config.json        # 中文格式配置
│   └── config_standard.json # 标准格式配置
├── examples/
│   └── sample_questions/  # 测试用示例数据
├── output/                # 生成的输出文件
├── run.py                # 用户友好界面
├── debug_excel.py        # Excel格式分析工具
├── main.py               # 交互模式
└── README.md             # 本文件
```

## 开发

### 测试

使用示例数据运行测试：
```bash
python create_samples.py
python src/merger.py --input examples/sample_questions
```

### 贡献

1. Fork 本仓库
2. 创建功能分支
3. 进行修改
4. 添加测试（如适用）
5. 提交Pull Request

## 许可证

本项目采用 MIT 许可证。详见 [LICENSE](LICENSE) 文件。

## 更新日志

### v1.0.0
- 初始版本，包含核心合并功能
- 支持多种Excel格式
- Word和Excel输出生成
- 智能格式检测
- 用户友好界面