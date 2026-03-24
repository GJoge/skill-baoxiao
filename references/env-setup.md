# 环境检查与安装指南

本文档提供 baoxiao 技能的完整环境检查和安装脚本。

## 快速检查

运行以下命令快速检查环境：

```bash
# 检查标记文件，若存在则跳过
if [ -f ~/.cache/baoxiao/.env_checked ]; then
    echo "✓ 环境已验证"
    exit 0
fi

# 执行完整检查
python3 ~/.claude/skills/baoxiao/scripts/check_env.py
```

## 依赖清单

### 系统依赖

| 依赖项 | 用途 | 检查命令 |
|--------|------|---------|
| tesseract-ocr | OCR识别 | `which tesseract` |
| tesseract-ocr-chi-sim | 中文识别 | `tesseract --list-langs \| grep chi_sim` |
| poppler-utils | PDF处理 | `which pdftotext` |
| libreoffice | 文档转换 | `which libreoffice` |
| ofd2pdf | OFD转换 | `which ofd2pdf` |

### Python依赖

```
pdfplumber>=0.10.0
pdf2image>=1.16.0
pytesseract>=0.3.10
Pillow>=9.0.0
openpyxl>=3.0.0
pandas>=1.3.0
python-docx>=0.8.11
reportlab>=3.6.0
pypdf>=3.0.0
pypdf2>=3.0.0
pyyaml>=6.0
```

### 字体依赖

识别中文发票需要安装常见字体：
- 宋体、楷体、黑体、仿宋
- 仿宋_GB2312、方正小标宋简体
- Arial

## 一键安装脚本

### Ubuntu/Debian

```bash
#!/bin/bash
set -e

echo "=== Baoxiao 环境安装 ==="

# 创建缓存目录
mkdir -p ~/.cache/baoxiao

# 安装系统依赖
echo "[1/3] 安装系统依赖..."
sudo apt-get update
sudo apt-get install -y \
    tesseract-ocr \
    tesseract-ocr-chi-sim \
    poppler-utils \
    libreoffice-writer \
    libreoffice-calc \
    fonts-wqy-zenhei \
    fonts-wqy-microhei

# 安装Python依赖
echo "[2/3] 安装Python依赖..."
pip3 install --user \
    pdfplumber pdf2image pytesseract pillow \
    openpyxl pandas python-docx reportlab \
    pypdf pypdf2 pyyaml

# 检查ofd2pdf
echo "[3/3] 检查OFD转换工具..."
if ! command -v ofd2pdf &> /dev/null; then
    echo "⚠️ ofd2pdf 未安装，OFD文件将无法转换"
    echo "  安装方法: https://github.com/Ofdcb/Ofd2Pdf"
fi

# 写入标记文件
touch ~/.cache/baoxiao/.env_checked
echo "✓ 环境检查完成"
```

### macOS

```bash
#!/bin/bash
set -e

echo "=== Baoxiao 环境安装 (macOS) ==="

# 创建缓存目录
mkdir -p ~/.cache/baoxiao

# 安装Homebrew依赖
echo "[1/3] 安装系统依赖..."
brew install tesseract tesseract-lang poppler libreoffice

# 安装Python依赖
echo "[2/3] 安装Python依赖..."
pip3 install --user \
    pdfplumber pdf2image pytesseract pillow \
    openpyxl pandas python-docx reportlab \
    pypdf pypdf2 pyyaml

# 写入标记文件
touch ~/.cache/baoxiao/.env_checked
echo "✓ 环境检查完成"
```

## 验证脚本

```bash
#!/bin/bash
# 保存为 check_env.sh

echo "=== Baoxiao 环境验证 ==="

errors=0

# 检查系统命令
check_cmd() {
    if command -v $1 &> /dev/null; then
        echo "✓ $1"
    else
        echo "✗ $1 未安装"
        ((errors++))
    fi
}

echo "[系统依赖]"
check_cmd tesseract
check_cmd pdftotext
check_cmd libreoffice

# 检查Python包
echo "[Python包]"
python3 -c "import pdfplumber, pytesseract, openpyxl, pandas, docx" 2>/dev/null && \
    echo "✓ Python依赖" || { echo "✗ Python依赖缺失"; ((errors++)); }

# 检查字体
echo "[字体]"
fc-list :lang=zh | grep -q "SimSun\|WenQuanYi" && \
    echo "✓ 中文字体" || { echo "⚠️ 中文字体可能缺失"; }

# 结果
if [ $errors -eq 0 ]; then
    echo ""
    echo "✓ 所有检查通过"
    mkdir -p ~/.cache/baoxiao
    touch ~/.cache/baoxiao/.env_checked
else
    echo ""
    echo "✗ 发现 $errors 个问题，请运行安装脚本"
    exit 1
fi
```

## 缓存机制说明

环境检查通过后会在 `~/.cache/baoxiao/.env_checked` 创建标记文件。

- **跳过检查**：标记文件存在时，直接执行主程序
- **强制重新检查**：删除标记文件 `rm ~/.cache/baoxiao/.env_checked`
- **缓存有效期**：永久有效（直到手动删除）

## 故障排除

### OCR识别中文失败

```bash
# 检查中文语言包
tesseract --list-langs | grep chi_sim

# 未安装则安装
sudo apt-get install tesseract-ocr-chi-sim  # Ubuntu
brew install tesseract-lang                 # macOS
```

### PDF转换失败

```bash
# 检查poppler
pdftotext -v

# 重新安装
sudo apt-get install --reinstall poppler-utils
```

### LibreOffice转换超时

大文件转换可能需要更长时间，可设置超时：

```bash
export LIBREOFFICE_TIMEOUT=300  # 5分钟
```