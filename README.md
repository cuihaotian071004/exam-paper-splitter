# 试卷分割工具 📄✂️

> 一键分离 DOCX/PDF 试卷中的题目与答案，保留图片/公式/表格，自动导出 PDF。

[![GitHub](https://img.shields.io/badge/GitHub-cuihaotian071004-181717?logo=github)](https://github.com/cuihaotian071004/exam-paper-splitter)
[![Version](https://img.shields.io/badge/version-3.1.0-orange)]()
[![Python](https://img.shields.io/badge/Python-3.11%2B-3776AB?logo=python)](https://python.org)
[![Windows](https://img.shields.io/badge/Windows-0078D6?logo=windows)]()

---

## 概述

**试卷分割工具**是一个 Windows 桌面工具，专为教育场景设计。它能够自动将 DOCX 或 PDF 格式的试卷按"参考答案"分界点，拆分为**题目**和**答案**两个独立的 PDF 文件。

依托 LibreOffice UNO API 实现高效 DOCX→PDF 转换，比传统方法快约 40 倍。

---

## 功能特性

| 特性 | 说明 |
|------|------|
| 🎯 **自动分割** | 按"参考答案"关键词自动定位分界点 |
| 🖼️ **完整保留** | 图片、公式、表格、排版全部保留 |
| ⚡ **高速转换** | 基于 LibreOffice UNO API 管道，DOCX→PDF 提速 40 倍 |
| 🔄 **降级保障** | UNO 失败时自动切换 soffice 直接转换 |
| 🖨️ **扫描 PDF** | 支持扫描型 PDF，提供 GUI 滑块手动选择分割页 |
| 🗑️ **空白页清理** | 自动删除输出 PDF 中的空白页 |
| 🖱️ **拖拽支持** | 支持文件拖拽到窗口直接处理 |
| 📦 **备份管理** | 源文件自动备份，支持一键打开备份文件夹 |
| 📋 **日志记录** | 每日日志文件，便于追踪处理历史 |

---

## 安装

### 前置依赖

- **Python 3.11+**
- **LibreOffice**（用于 DOCX→PDF 转换）
  - 下载: [LibreOffice 官网](https://www.libreoffice.org/download/)
  - 安装时选择"安装到本机"即可
- **Python 依赖包**

```bash
pip install PyMuPDF tkinterdnd2
```

> `tkinterdnd2` 为可选项（提供拖拽支持），如不安装工具仍可正常使用。

### 快速启动

双击 `试卷分割工具.bat` 或 `试卷分割工具.pyw` 即可启动。

首次启动会自动检测 LibreOffice 并初始化转换引擎。

---

## 使用指南

### 基本流程

1. **启动工具** → 自动初始化 LibreOffice 引擎
2. **选择文件** — 点击"选择文件"或直接拖拽 DOCX/PDF 到窗口
3. **自动处理** — 工具自动完成转换 → 分割 → 导出
4. **获取结果** — 在源文件所在文件夹得到 `_题目.pdf` 和 `_答案.pdf`

### 文件支持

| 格式 | 处理流程 |
|------|----------|
| **DOCX** | DOCX → PDF (UNO API) → 分割 → 输出 |
| **PDF** | PDF → 直接分割 → 输出 |
| **扫描 PDF** | PDF → 手动选择分割页 → 输出 |

### 扫描 PDF 处理

对于扫描型 PDF（无文字层），工具会弹出 GUI 窗口：
1. 拖动滑块预览各页内容
2. 选择"参考答案"开始的分界页
3. 点击"确认分割"

---

## 输出文件

| 文件 | 说明 |
|------|------|
| `原文件名_题目.pdf` | 仅包含题目部分 |
| `原文件名_答案.pdf` | 仅包含答案部分 |

源文件自动移动到备份目录：`%LOCALAPPDATA%\Hermes Agent CN Desktop\docx_backup\`

---

## 项目结构

```
exam-paper-splitter/
├── 试卷分割工具.pyw    # GUI 主程序 (Tkinter)
├── core.py             # 核心处理逻辑 (v3.1.1)
├── lo_convert.py       # LibreOffice UNO 转换脚本
├── 试卷分割工具.bat     # 启动脚本
├── .gitignore
├── logs/               # 日志目录
└── README.md
```

---

## 技术原理

1. **DOCX→PDF 转换**: 通过 UNO socket 连接 LibreOffice 监听器，调用 `writer_pdf_Export` 过滤器直接导出 PDF
2. **分割定位**: 逐页扫描 PDF 文本内容，查找"参考答案"或独立的"答案"标题作为分界点
3. **PDF 操作**: 使用 PyMuPDF (fitz) 进行页面提取、合并和空白页检测

---

## English Version

### Overview

**Exam Paper Splitter** is a Windows desktop tool for educators. It automatically separates exam papers into **questions** and **answers** as two independent PDF files, using the "参考答案" (reference answer) marker as the split point.

### Features

- Auto-detect answer section boundary
- Preserve images, formulas, and tables
- LibreOffice UNO API pipeline for fast DOCX→PDF
- Scanned PDF support with manual page selector GUI
- Auto-cleanup of empty pages
- Drag-and-drop support
- Backup management

### Quick Start

```bash
# Install dependencies
pip install PyMuPDF tkinterdnd2

# Run
python "试卷分割工具.pyw"
```

---

## 许可证

MIT License

---

**Made with ❤️ for Chinese educators / 为中国教育工作者制作**
