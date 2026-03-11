# QCT_Tools
A small tool that automatically generates QCT files and Comments files from PDT files.
一个简单易用的 **PDT 文件转 QCT 文件和审阅意见文件的工具**，帮助快速处理文档，无需安装 Python 环境。

---
## 📋 项目背景
在项目运行过程中，程序员需要用到QCT文件，记录QC的findings和解决方案。
当项目交付时，combined的pdf文件过大不便添加批注时，项目程序员则需要提供审阅意见文件给到项目组，用于添加审阅意见。
传统方式下，程序员需要手动填写这俩份文档，效率低且不规范。**QCT_Tools**可以读取PDT表格的信息，自动提取表格编号、标题、开发人员、QC人员，大幅减少手动输入的工作。

## ✨ 功能特点

- ✅ **免安装**：解压即用，无需配置 Python 环境
- ✅ **直观操作**：图形界面，点击鼠标就能完成
- ✅ **叠加处理**：QCT文件支持不同 Event 叠加
- ✅ **双文件导出**：支持导出 QCT 文件和审阅意见文件

---

## 🚀 快速开始

### 系统要求
- Windows 7/10/11 (64位)
- 无需安装 Python 或其他依赖

### 下载与使用
1. **下载工具**  
   前往 [Releases 页面]https://github.com/mochao666/QCT_Tools/releases 下载最新版本的 `QCT_Tools.zip`

2. **解压文件**  
   将压缩包解压到任意文件夹  
   ⚠️ **重要提示：解压后请保留文件夹内所有文件，不要只复制 exe 文件**

3. **运行工具**  
   双击 `QCT_Tools.exe` 即可打开工具界面

---

## 📖 使用说明

### 操作步骤

1. **导入数据**  
   双击运行 QCT_Tools.exe 打开小工具。点击「导入 PDT」按钮，选择你的 PDT Excel 文件，可以选择对应的Event

2. **导出结果**  
   - 点击「导出 QCT」：保存为 QCT 格式文件
   - 点击「导出 Comments」：保存为审阅意见文件
   - comments 文件会保留F 列（Developers） 和 G 列（Validators），并设为隐藏，在 Excel 里默认不显示。

3. **不同Event叠加**  
   - 首次导出QCT文件可以选「初版 QCT 」
   - 在一轮QC完成后，新增Event进行新一轮的QC，可以重新导入 PDT → 将 Event 改为另一个Event → 导出QCT → 选「新增Event」→ 会叠加新的行，同时将原来QCT中的空行删除，仅保留有QC comments的行
   - 终版QCT：在项目结束后，可以导入QCT → 导出 QCT → 选「终版 QCT 」，会删除QCT中的空行，重新保存。
     
4. **权限控制**  
   - 对于没有权限的文件夹只能导入PDT，无法导出QCT。会显示Permission denied.
