# 部署指南

## GitHub 部署步骤

### 1. 创建 GitHub 仓库

```bash
git init
git add .
git commit -m "Initial commit: 题库合并工具 v1.0"

# 关联到你的GitHub仓库
git remote add origin https://github.com/your-username/question-bank-merger.git
git push -u origin main
```

### 2. 设置 GitHub Pages（可选）

在仓库设置中启用 GitHub Pages，选择 main 分支的 /docs 目录或 root 目录，这样你的 README.md 就可以在线访问。

### 3. 创建 Release

```bash
git tag v1.0.0
git push origin v1.0.0
```

### 4. 添加 GitHub Actions（可选）

创建 `.github/workflows/test.yml`:

```yaml
name: Test

on: [push, pull_request]

jobs:
  test:
    runs-on: ubuntu-latest
    strategy:
      matrix:
        python-version: [3.8, 3.9, '3.10']

    steps:
    - uses: actions/checkout@v2
    - name: Set up Python
      uses: actions/setup-python@v2
      with:
        python-version: ${{ matrix.python-version }}
    - name: Install dependencies
      run: |
        pip install -r requirements.txt
    - name: Test with sample data
      run: |
        python create_samples.py
        python src/merger.py --input examples/sample_questions
```

## 使用说明（面向用户）

### 安装

```bash
# 克隆仓库
git clone https://github.com/your-username/question-bank-merger.git
cd question-bank-merger

# 安装依赖
pip install -r requirements.txt
```

### 基本使用

1. **将你的Excel题库文件放在项目目录中**

2. **运行合并工具**
   ```bash
   # 使用默认配置（中文题库格式）
   python src/merger.py

   # 或使用主程序
   python main.py
   ```

3. **查看输出**
   - Excel文件：`output/merged_questions.xlsx`
   - Word文件：`output/merged_questions.docx`

### 高级使用

```bash
# 自定义配置
python src/merger.py --config config/my_config.json

# 指定输入目录
python src/merger.py --input /path/to/questions

# 只生成Word文档
python src/merger.py --word-only
```

## 贡献指南

欢迎提交 Issue 和 Pull Request！

1. Fork 项目
2. 创建功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启 Pull Request

## 许可证

MIT License - 详见 LICENSE 文件