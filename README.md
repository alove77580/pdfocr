# PDF OCR识别工具

这是一个基于Python的Windows可视化应用程序，用于对PDF文件进行OCR文字识别。

## 功能特点

- 支持选择本地PDF文件
- 支持中英文混合识别
- 显示识别进度
- 实时显示识别结果

## 安装要求

1. Python 3.7+
2. Tesseract-OCR
3. Poppler

## 安装步骤

1. 安装Tesseract-OCR
   - 从 https://github.com/UB-Mannheim/tesseract/wiki 下载并安装
   - 确保将Tesseract添加到系统环境变量中

2. 安装Poppler
   - 从 http://blog.alivate.com.au/poppler-windows/ 下载
   - 解压并将bin目录添加到系统环境变量中

3. 安装Python依赖
```bash
pip install -r requirements.txt
```

## 使用方法

1. 运行程序：
```bash
python main.py
```

2. 点击"选择PDF文件"按钮
3. 选择要识别的PDF文件
4. 等待识别完成
5. 查看识别结果

## 注意事项

- 确保PDF文件清晰可读
- 识别过程可能需要一些时间，请耐心等待
- 如果遇到错误，请检查Tesseract和Poppler是否正确安装 