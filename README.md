
```markdown
# 假条生成器 (Leave Application Generator)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python Version](https://img.shields.io/badge/Python-3.6%2B-blue.svg)](https://www.python.org/)

假条生成器项目是一个基于 Python 和 Tkinter 的桌面应用程序，用于批量生成请假条文档。项目支持 Excel 表格输入和文本输入两种方式，用户可以根据需求灵活生成假条。生成的假条会自动保存到桌面，方便快速打印或分发。此外，本项目还提供了 PyInstaller 打包工具，可以将项目打包为单文件可执行文件，便于在没有 Python 环境的电脑上运行。

---

## 目录结构

```
leave_application_project/
├── README.md              # 本文档，介绍项目详情和使用说明
├── requirements.txt       # 项目所需依赖库列表
├── setup.py               # 可选的安装脚本（发布或安装时使用）
├── src/
│   ├── __init__.py        # 包初始化文件
│   ├── main.py            # 程序入口，启动 GUI 界面
│   ├── gui.py             # 图形用户界面模块，负责数据交互与用户输入
│   ├── document_generator.py  # 文档生成模块，根据模板和数据生成假条文档
│   └── package_creator.py       # 打包工具模块，使用 PyInstaller 打包项目
└── templates/
    └── 请假条模板.docx     # 假条模板文件，定义了假条样式与占位符
```

---

## 项目背景与功能

### 背景介绍

在日常教学、工作或生活中，经常需要编写各类请假条。手动填写假条不仅效率低，而且容易出错。本项目通过自动化的方式，从用户提供的数据中批量生成标准格式的假条文档，既保证格式统一，也大幅提升工作效率。

### 主要功能

- **多种数据输入方式**
  - **Excel 表格输入**：支持批量导入假条信息。
  - **文本输入**：通过交互式窗口依次录入各个假条的信息，适合人数较少时使用。

- **模板替换机制**
  - 根据预设的请假条模板（`请假条模板.docx`），自动替换文档中的占位符（如 `input1`、`input2` 等）为用户输入的数据，确保生成的文档格式标准、内容准确。

- **简洁直观的 GUI 界面**
  - 基于 Tkinter 设计，界面清晰易用，支持生成空表格模板、Excel 文件导入、以及文本逐条输入假条信息。

- **打包工具支持**
  - 通过 PyInstaller 打包工具，可将项目打包成单个可执行文件，使用户无需 Python 环境即可运行程序。

---

## 安装与依赖

本项目建议使用 Python 3.6 及以上版本。项目依赖如下：

- [pandas](https://pandas.pydata.org/)
- [python-docx](https://python-docx.readthedocs.io/)
- [openpyxl](https://openpyxl.readthedocs.io/)
- [pyinstaller](https://www.pyinstaller.org/)

安装所有依赖：

```bash
pip install -r requirements.txt
```

---

## 使用说明

### 1. 运行程序

进入 `src` 目录后，执行以下命令启动程序：

```bash
python main.py
```

启动后，图形界面会提供以下三种选择：

- **表格输入**  
  导入包含假条信息的 Excel 文件，程序自动读取并生成假条。

- **文本输入**  
  通过交互式窗口逐条录入假条信息。

- **生成模板表格**  
  自动生成符合要求的空白 Excel 表格模板，供用户填写后导入。

### 2. 打包程序

若需将项目打包为单文件可执行文件，可使用以下命令：

```bash
python src/package_creator.py
```

打包成功后，生成的可执行文件将保存在项目根目录下的 `dist` 文件夹内。

---

## 模块说明

- **main.py**  
  程序入口，负责启动图形用户界面。

- **gui.py**  
  实现图形界面，封装所有用户交互逻辑，如数据录入、文件选择、表格生成等。

- **document_generator.py**  
  根据模板文件和用户输入数据生成最终的假条文档，并将生成结果保存到桌面。

- **package_creator.py**  
  使用 PyInstaller 将项目打包成单文件可执行文件，自动处理模板文件的包含与路径配置。

---

## 模板文件说明

模板文件位于 `templates/请假条模板.docx`，定义了假条的基本格式和占位符，包括：

- `input1`：学院  
- `input2`：专业  
- `input3`：班级  
- `input4`：姓名  
- `input5`：学号  
- `input6`：事由  
- `input7`：请假时间  
- `y`、`m`、`d`：日期（年、月、日）

程序将自动替换这些占位符为用户提供的信息。

---

## 贡献与维护

欢迎各位同志参与贡献代码、提交改进建议或报告问题！

- **Fork 项目** 后提交 Pull Request
- **提交 Issue** 反馈问题或讨论改进方案

请确保提交代码遵循 [PEP8](https://www.python.org/dev/peps/pep-0008/) 代码风格，并附带必要的注释说明。

---

## 许可证

本项目采用 [MIT 许可证](LICENSE) 开源，任何人均可自由使用、修改和分发代码，但请保留原作者信息。

---

## 常见问题 (FAQ)

**Q：运行程序时提示找不到模板文件怎么办？**  
A：请确保 `请假条模板.docx` 文件位于 `templates` 文件夹中，并且目录结构未被更改。

**Q：如何生成空白 Excel 模板？**  
A：点击主界面中的“生成模板表格”按钮，系统会自动生成符合要求的 Excel 模板。

**Q：如何打包程序？**  
A：安装 `pyinstaller` 后，运行 `python src/package_creator.py` 即可生成单文件可执行程序。

---

## 更新日志

### v1.0.0

- 初始版本：实现 Excel 表格输入、文本输入两种数据输入方式
- 自动替换请假条模板占位符生成标准假条文档
- 提供 PyInstaller 打包工具，生成单文件可执行程序

---
