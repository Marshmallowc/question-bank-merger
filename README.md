# Question Bank Merger

A powerful tool for merging multiple Excel question banks into unified documents. Designed for educators, trainers, and educational institutions who need to consolidate scattered question banks.

## Overview

Question Bank Merger automatically detects and merges Excel files containing questions, supporting various formats and generating both Word and Excel outputs. No programming knowledge required.

## Features

- **Smart Detection**: Automatically identifies Excel file formats and column structures
- **One-Click Operation**: Simply run the tool and let it handle everything
- **Multiple Output Formats**: Generate both Excel and Word documents
- **Flexible Configuration**: Support for various question bank formats through JSON configuration
- **Error Handling**: Comprehensive error reporting and troubleshooting guidance
- **Chinese Language Support**: Optimized for Chinese educational question formats

## Quick Start

### Prerequisites

- Python 3.7 or higher
- Required packages: pandas, openpyxl, python-docx

### Installation

1. Clone the repository
   ```bash
   git clone https://github.com/Marshmallowc/question-bank-merger.git
   cd question-bank-merger
   ```

2. Install dependencies
   ```bash
   pip install pandas openpyxl python-docx
   ```

3. Prepare your Excel question bank files in the project directory

4. Run the merger
   ```bash
   python run.py
   ```

## Usage

### Basic Usage (Recommended for Beginners)

The tool provides an interactive interface that guides you through the process:

1. Place your Excel files in the project directory
2. Run `python run.py`
3. Follow the on-screen instructions to select files
4. Generated files will be saved in the `output/` directory

### Advanced Usage

For more control over the merging process:

```bash
# Use default configuration
python src/merger.py

# Specify custom configuration
python src/merger.py --config config/my_config.json

# Specify input directory and file pattern
python src/merger.py --input /path/to/questions --pattern "chapter*.xlsx"

# Generate only Word output
python src/merger.py --word-only

# Generate only Excel output
python src/merger.py --excel-only
```

## Configuration

The tool supports various Excel formats through JSON configuration files. Default configurations are included:

- `config/config.json` - Chinese question bank format (with header description row)
- `config/config_standard.json` - Standard Excel format

### Custom Configuration

Create a custom JSON configuration file to match your specific format:

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

## Supported Question Formats

- Multiple Choice Questions
- True/False Questions
- Fill in the Blanks
- Short Answer Questions
- Essay Questions
- Custom question types

## Troubleshooting

### Common Issues

**Files not detected**
- Ensure Excel files are in the correct directory
- Check file extensions (.xlsx or .xls)
- Verify files are not password-protected

**Merge failures**
- Try resaving Excel files in .xlsx format
- Check for corrupted files
- Use the debug tool: `python debug_excel.py your_file.xlsx`

**Dependency installation errors**
- Update pip: `python -m pip install --upgrade pip`
- Use alternative pip source if needed:
  ```bash
  pip install -i https://pypi.tuna.tsinghua.edu.cn/simple/ pandas openpyxl python-docx
  ```

## Examples

### Example 1: University Question Banks

Your Excel files might look like this:

| 题型 | 题干 | 正确答案 | 选项A | 选项B | 选项C | 选项D |
|------|------|----------|-------|-------|-------|-------|
| 单选题 | 下列哪个是Python的特点？ | B | 编译型语言 | 解释型语言 | 汇编语言 | 机器语言 |

The tool will automatically detect this format and merge multiple such files.

### Example 2: English Questions

| Question Type | Question | Answer | Option A | Option B | Option C | Option D |
|---------------|----------|--------|----------|----------|----------|----------|
| Multiple Choice | What is 2+2? | B | 3 | 4 | 5 | 6 |

## Project Structure

```
question-bank-merger/
├── src/
│   └── merger.py          # Core merging logic
├── config/
│   ├── config.json        # Chinese format configuration
│   └── config_standard.json # Standard format configuration
├── examples/
│   └── sample_questions/  # Sample data for testing
├── output/                # Generated output files
├── run.py                # User-friendly interface
├── debug_excel.py        # Excel format analysis tool
├── main.py               # Interactive mode
└── README.md             # This file
```

## Development

### Testing

Run the test suite with sample data:
```bash
python create_samples.py
python src/merger.py --input examples/sample_questions
```

### Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.

## Changelog

### v1.0.0
- Initial release with core merging functionality
- Support for multiple Excel formats
- Word and Excel output generation
- Smart format detection
- User-friendly interface