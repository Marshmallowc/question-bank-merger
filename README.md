# 题库合并工具 (Question Bank Merger)

一个通用的题库合并工具，可以将多个Excel格式的题库文件合并成一个统一的Word文档或Excel文件。该工具特别适合教育机构、培训中心、教师等需要整理分散题库的用户。

## 功能特点

- **自动合并**：自动合并目录中的所有题库Excel文件
- **多格式输出**：支持生成Excel和Word两种格式
- **灵活配置**：通过配置文件适应不同的Excel格式
- **智能检测**：自动识别表头位置和数据结构
- **统计报告**：生成详细的合并统计信息
- **中文支持**：完美支持中文内容和格式

## 安装依赖

```bash
pip install pandas openpyxl python-docx
```

## 快速开始

### 1. 克隆项目

```bash
git clone https://github.com/your-username/question-bank-merger.git
cd question-bank-merger
```

### 2. 准备题库文件

将你的Excel题库文件放在项目目录中，支持的模式：
- `*_习题导出.xlsx`
- `*questions*.xlsx`
- `*题库*.xlsx`
- `*.xlsx`

### 3. 运行合并

```bash
# 使用默认配置
python src/merger.py

# 指定配置文件
python src/merger.py --config config/config_standard.json

# 指定输入目录和文件模式
python src/merger.py --input /path/to/questions --pattern "*题库*.xlsx"
```

### 4. 查看输出

合并后的文件将保存在 `output/` 目录中：
- `merged_questions.xlsx` - Excel格式
- `merged_questions.docx` - Word格式

## 配置说明

项目包含两个预设配置：

### 1. config.json (默认配置)
适合中国大学的题库格式：
- 第一行是说明文字
- 第二行是列名（题型、题干、正确答案等）
- 第三行开始是数据

### 2. config_standard.json (标准格式)
适合标准Excel格式：
- 第一行是列名
- 第二行开始是数据

### 自定义配置

你可以创建自己的配置文件，示例：

```json
{
  "excel_settings": {
    "has_header_row": true,
    "header_row_index": 0,
    "data_start_row": 1,
    "skip_description_row": false
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
    "excel_filename": "output/my_questions.xlsx",
    "word_filename": "output/my_questions.docx",
    "include_analysis": true,
    "include_difficulty": true
  },
  "file_patterns": [
    "*.xlsx"
  ]
}
```

## 命令行参数

```bash
python src/merger.py [选项]

选项：
  --config PATH     配置文件路径 (默认: config/config.json)
  --input PATH      输入目录 (默认: 当前目录)
  --pattern PATTERN 文件匹配模式
  --output-excel    Excel输出文件路径
  --output-word     Word输出文件路径
  --word-only       只生成Word文档
  --excel-only      只生成Excel文件
```

## 使用场景与配置示例

### 场景一：标准中文题库格式

假设你的Excel文件具有以下格式：

| 题型 | 题干 | 正确答案 | 解析 | 分值 | 难度系数 | 选项A | 选项B | 选项C | 选项D | 选项E |
|------|------|----------|------|------|----------|-------|-------|-------|-------|-------|
| 单选题 | 下列哪个是Python的特点？ | B | Python是解释型语言 | 1.0 | 1 | 编译型语言 | 解释型语言 | 汇编语言 | 机器语言 | |
| 多选题 | Python支持哪些数据类型？ | ABCD | Python支持多种数据类型 | 2.0 | 2 | 整数 | 字符串 | 列表 | 字典 | 集合 |

**使用默认配置即可**：
```bash
python src/merger.py
```

### 场景二：英文标准格式

假设你的Excel文件具有英文列名：

| Question Type | Question | Answer | Analysis | Score | Option A | Option B | Option C | Option D |
|---------------|----------|--------|----------|-------|----------|----------|----------|----------|
| Multiple Choice | What is Python? | B | Python is an interpreted language | 1 | Compiled language | Interpreted language | Assembly language | Machine language |

**使用标准配置**：
```bash
python src/merger.py --config config/config_standard.json
```

### 场景三：自定义列名

如果你的Excel文件使用了不同的列名，例如：

| 问题类型 | 题目内容 | 答案 | 选项1 | 选项2 | 选项3 | 选项4 |
|----------|----------|------|-------|-------|-------|-------|
| 选择题 | 首都是哪里？ | A | 北京 | 上海 | 广州 | 深圳 |

**需要创建自定义配置文件** `config/my_config.json`：

```json
{
  "excel_settings": {
    "has_header_row": true,
    "header_row_index": 0,
    "data_start_row": 1,
    "skip_description_row": false
  },
  "column_mapping": {
    "question_type": "问题类型",
    "question_text": "题目内容",
    "correct_answer": "答案",
    "analysis": "解析",
    "score": "分值",
    "difficulty": "难度",
    "options": ["选项1", "选项2", "选项3", "选项4"]
  },
  "output_settings": {
    "excel_filename": "output/custom_merged.xlsx",
    "word_filename": "output/custom_merged.docx",
    "include_analysis": false,
    "include_difficulty": false
  },
  "file_patterns": [
    "*.xlsx",
    "*题库*.xlsx"
  ]
}
```

**使用自定义配置**：
```bash
python src/merger.py --config config/my_config.json
```

## 调试指南

### 常见问题及解决方案

#### 1. 列名不匹配错误

**错误信息**：
```
KeyError: '题型'
```

**原因**：配置文件中的列名与实际Excel列名不匹配。

**解决步骤**：

1. **使用调试工具分析Excel文件**：
   ```bash
   python debug_excel.py your_file.xlsx
   ```
   这个工具会自动分析你的Excel文件格式，并推荐合适的配置。

2. **或者手动查看列名**：
   ```python
   import pandas as pd
   df = pd.read_excel('your_file.xlsx', engine='openpyxl')
   print(list(df.columns))
   ```

3. **修改配置文件**，将 `column_mapping` 中的值改为实际的列名。

4. **如果第一行是说明文字**，确保配置正确：
   ```json
   "excel_settings": {
     "skip_description_row": true,
     "description_row_index": 0,
     "header_row_index": 1
   }
   ```

#### 2. 答案列为空

**可能原因**：
- 答案列的名称配置错误
- 答案数据在Excel中确实为空

**解决方案**：
1. 确认配置中的 `"correct_answer"` 值与Excel中的答案列名完全一致
2. 检查Excel文件中答案列是否有数据

#### 3. 多个文件读取失败

**错误信息**：
```
找到 5 个文件
正在读取: file1.xlsx
  ✗ 读取失败: ...
```

**可能原因**：
- 文件格式不是标准的xlsx格式
- 文件损坏
- Excel文件有特殊的保护

**解决方案**：
1. 用Excel打开文件，另存为新的xlsx文件
2. 确保文件没有被密码保护
3. 检查文件是否可以正常打开

#### 4. 输出文件格式错误

**问题**：生成的Word或Excel文件格式异常

**解决步骤**：
1. 确保已安装所有依赖：
   ```bash
   pip install pandas openpyxl python-docx
   ```
2. 检查输出目录权限
3. 尝试使用绝对路径作为输出路径

### 使用示例

### 示例1：合并当前目录的所有题库文件

```bash
python src/merger.py
```

### 示例2：合并指定目录的特定格式文件

```bash
python src/merger.py --input /path/to/excel --pattern "chapter*.xlsx"
```

### 示例3：只生成Word文档

```bash
python src/merger.py --word-only
```

### 示例4：使用自定义配置和输出路径

```bash
python src/merger.py \
  --config config/my_config.json \
  --output-excel output/biology_questions.xlsx \
  --output-word output/biology_questions.docx
```

## 项目结构

```
question-bank-merger/
├── src/
│   └── merger.py          # 主程序
├── config/
│   ├── config.json        # 默认配置（中文题库格式）
│   └── config_standard.json # 标准配置
├── examples/
│   └── sample_questions/  # 示例数据
├── output/                # 输出目录
├── README.md              # 说明文档
├── requirements.txt       # 依赖列表
└── .gitignore            # Git忽略文件
```

## 支持的题型

- 单选题
- 多选题
- 判断题
- 填空题
- 主观题
- 其他自定义题型

## 贡献

欢迎提交Issue和Pull Request来改进这个工具！

## 许可证

MIT License - 详见 LICENSE 文件

## Star History

如果这个项目对你有帮助，请给个Star支持一下！