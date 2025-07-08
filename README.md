# Word to PDF Converter

## 📌项目简介

这是一个 Python 脚本，用于将 `.docx` 文件转换为 `.pdf` 文件。支持单个文件转换或批量转换整个目录下的所有 `.docx` 文件。

## 功能

* 将单个`.docx`文件转换为`.pdf`文件
* 批量转换指定目录下的所有 `.docx` 文件
* 错误处理：自动跳过非 `.docx` 文件，并显示转换失败的原因

## 👀️ 安装依赖

在使用脚本之前请确认已安装Python3和以下库：

```
pip install docx2pdf argparse
```

## 🚀️ 使用方法

### 1.转换单个文件：

```
python .\word2pdf.py [文件路径.docx]
```

示例：

```
python .\word2pdf.py "C:/Documents/example.docx"
```

### 2. 批量转换目录下的所有 `.docx` 文件:

```
python .\word2pdf.py [目录路径]
```

示例：

```
python .\word2pdf.py "C:\Documents\pWordFiles"
```

## 📝 注意事项

1. 仅支持 `.docx` 格式，不支持旧版 `.doc` 文件。
2. 转换后的 PDF 文件会保存在与原文件相同的目录下。
3. 如果路径包含空格或特殊字符，请使用引号包裹路径（如 `"path with spaces"`）。

## ❓ 常见问题

### Q: 转换失败怎么办？

* 确保输入的文件是有效的 `.docx` 文件。
* 检查文件是否被其他程序（如 Word）占用。
* 查看错误信息，确认是否有权限问题。
