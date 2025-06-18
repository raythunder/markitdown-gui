# MarkItDown GUI

一个基于 [Microsoft MarkItDown](https://github.com/microsoft/markitdown) 的图形界面工具，可以将各种文档格式转换为 Markdown。

## ✨ 功能特性

### 🔄 **多格式支持**
- **Office 文档**: Word (.docx), Excel (.xlsx), PowerPoint (.pptx)
- **PDF 文档**: 支持文本和图片提取
- **电子书**: EPUB 格式，支持图片提取
- **音频文件**: MP3, WAV, M4A 等（通过 AI 转录）
- **图片文件**: JPG, PNG, GIF 等（通过 OCR 识别）
- **网页文件**: HTML
- **其他格式**: CSV, JSON, XML 等

### 📚 **EPUB 图片处理**
- ✅ **自动图片提取**: 从 EPUB 文件中提取所有图片
- ✅ **智能路径处理**: 支持相对路径、绝对路径、Base64 编码
- ✅ **完整保存**: 保存时自动将图片复制到输出目录
- ✅ **格式支持**: JPG、PNG、GIF、BMP、SVG、WebP

### 🎯 **用户体验**
- 🖱️ **简洁界面**: 直观的图形用户界面
- ⚡ **多线程处理**: 转换过程不阻塞界面
- 📝 **实时预览**: 转换结果即时显示
- 💾 **灵活保存**: 支持另存为功能
- 🔧 **智能配置**: 记住用户偏好设置

## 📦 安装要求

### 系统要求
- Windows 10/11
- Python 3.8+

### 依赖安装
```bash
# 安装 MarkItDown 及所有依赖
pip install markitdown[all]

# 安装 GUI 依赖
pip install tkinter  # 通常已内置于 Python
```

## 🚀 快速开始

### 直接运行
```bash
python markitdown_gui.py
```

### 或者构建可执行文件
```bash
python build_exe.py
```

## 📖 使用指南

### 基本转换流程
1. **选择文件**: 点击"浏览..."选择要转换的文件
2. **配置选项**: 
   - 对于 EPUB 文件，可选择是否提取图片
   - 选择图片引用方式（相对路径/绝对路径/Base64）
3. **开始转换**: 点击"转换为Markdown"
4. **查看结果**: 在结果区域预览转换后的内容
5. **保存文件**: 点击"保存结果"选择保存位置

### EPUB 图片处理
转换 EPUB 文件时：
- ✅ 勾选"从EPUB提取图片"自动提取图片文件
- 📁 图片将保存在输出目录的 `images/` 文件夹中
- 🔗 Markdown 中的图片引用会自动更新为正确路径
- 💾 保存时图片会自动复制到目标目录

## 🏗️ 项目结构

```
markitdown-gui/
├── markitdown_gui.py     # 主程序文件
├── build_exe.py          # 可执行文件构建脚本
├── build.bat            # Windows 构建脚本
├── test_gui.py          # GUI 测试脚本
├── icon.ico             # 程序图标
├── README.md            # 项目说明
└── 使用说明.md          # 中文使用说明
```

## 🛠️ 开发说明

### 核心技术栈
- **GUI 框架**: Tkinter
- **转换引擎**: Microsoft MarkItDown
- **图片处理**: PIL/Pillow
- **文件处理**: zipfile, os, shutil
- **多线程**: threading

### 主要功能模块
- `MarkItDownGUI`: 主界面类
- `extract_epub_images()`: EPUB 图片提取
- `process_markdown_images()`: 图片引用处理
- `copy_images_to_target()`: 图片文件复制

## 🐛 故障排除

### 常见问题
1. **音频转换失败**: 需要安装 FFmpeg
2. **某些 PDF 无法转换**: 可能是加密或特殊格式
3. **EPUB 图片不显示**: 确保勾选了"从EPUB提取图片"选项

### 解决方案
```bash
# 安装 FFmpeg (用于音频处理)
# Windows: 下载并添加到 PATH
# 或使用 chocolatey: choco install ffmpeg

# 更新依赖
pip install --upgrade markitdown[all]
```

## 📄 许可证

本项目基于 MIT 许可证开源。

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

## 📞 联系方式

- GitHub: [raythunder/markitdown-gui](https://github.com/raythunder/markitdown-gui)
- Issues: [提交问题](https://github.com/raythunder/markitdown-gui/issues)

---

**感谢使用 MarkItDown GUI！** 🎉 