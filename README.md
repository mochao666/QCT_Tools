# QCT_Tools

基于 Python 的 **PDT 与 QCT 联动** 小工具：从 PDT 文件读取任务分配信息，自动生成带预填内容的 QCT 文件，无需每个 study 重新建 QCT、重复填写 SDTM/ADaM/TFL 及开发/QC 人员。

## 环境要求

- Python 3.8+
- 依赖：`pandas`、`openpyxl`

## 安装依赖

```bash
pip install -r requirements.txt
```

## 功能说明

1. **读取 PDT**：用 `pandas.read_excel` 读取 PDT，提取列组成“主词典”。
2. **所需列**：`Output Type`, `Output Reference`, `Title`, `Developers`, `Validators`, `Date Checked by Trial Statistician`。
3. **生成 QCT**：
   - **Output Type = "SDTM"** → 写入 Sheet **SDTM(aCRF, SPEC and Coding)**  
     列映射：Output Reference → SDTM Datasets，Title → QC checklist-index，Developers → Person Responsible for Resolution if Findings，Validators → QC programmer Name，Date Checked by Trial Statistician → Date of QC。
   - **其他类型（ADaM / TFL）** → 写入 Sheet **ADaM(SPEC and Coding) and TFL**  
     列映射：Output Reference → ADaM Dataset / TFL Number，其余同上。
   - `QC results description`、`Specify Notes (If Final Status="Followup")`、`Final Status` 等留空，由后续人工填写。
4. **输出**：保存为新 Excel，默认文件名可带 `_QCT_Template`，也可自定义（如 `HRxxxxx_xxx_CSR_01_QCT_v1.0_Template.xlsx`）。

## 使用方法

### 小工具界面（推荐）

在工具内导入 PDT、生成 QCT 模板、填写 **QC results description** 和 **Specify Notes (If Final Status="Followup")**，再导出 QCT：

```bash
python app_gui.py
```

或双击运行 **`qct_app.bat`**（Windows，需在 QCT_Tools 文件夹内）。

**若要在桌面使用**：在 QCT_Tools 文件夹内双击运行 **`创建桌面快捷方式.bat`** 一次，会在桌面生成可用的 `qct_app.bat`，之后在桌面双击即可打开小工具。

1. 点击 **「导入 PDT」**，选择 PDT 的 Excel 文件。
2. 工具会按 Sheet 生成 QCT 表格（SDTM / ADaM&TFL），可在下方 **「当前行编辑」** 中填写 QC results description 与 Specify Notes。
3. 在表格中选中某一行，在底部编辑框内填写内容，点击 **「应用到此行」** 保存到当前行。
4. 点击 **「导出 QCT」**，选择保存路径即可得到完整 QCT 文件。

### 命令行：从 PDT 直接生成 QCT

```bash
# 指定 PDT 文件，输出到默认路径（同目录下 原文件名_QCT_Template.xlsx）
python pdt_to_qct.py "C:\path\to\PDT_Study123.xlsx"

# 指定输出 QCT 路径
python pdt_to_qct.py "PDT_Study123.xlsx" -o "HRxxxxx_xxx_CSR_01_QCT_v1.0_Template.xlsx"

# 使用自定义 QCT 模板（仅表头）
python pdt_to_qct.py "PDT_Study123.xlsx" -o "QCT_out.xlsx" -t "QCT_template.xlsx"
```

### 仅生成 QCT 表头模板（可选）

若需先得到“只有表头”的模板再手工调整列名或顺序，可运行：

```bash
python create_qct_template.py -o "QCT_template.xlsx"
```

之后用 `pdt_to_qct.py -t QCT_template.xlsx` 指定该模板。

## 打包成可执行程序（上传/分发用）

可将 QCT_Tools 打成 **免安装** 的 Windows 可执行程序，任意电脑下载后解压即可使用，**无需安装 Python**。

### 打包步骤（在已安装 Python 的本机执行一次）

1. 安装打包依赖：
   ```bash
   pip install -r requirements-dev.txt
   ```
2. 在项目目录下双击运行 **`build.bat`**（或命令行执行 `pyinstaller --noconfirm --clean QCT_Tools.spec`）。
3. 打包完成后，在 **`dist\QCT_Tools`** 下会生成整个程序文件夹，其中 **`QCT_Tools.exe`** 为主程序。

### 分发与使用（接收方电脑）

1. 将 **`dist\QCT_Tools`** 整个文件夹压缩成 **zip**（或 7z），上传到网盘/内网/邮件等。
2. 其他电脑：**下载 zip → 解压到任意位置**（如桌面、D 盘）。
3. 双击解压后的 **`QCT_Tools.exe`** 即可打开 QCT 小工具，无需安装 Python 或任何依赖。
4. 使用方式与界面版一致：导入 PDT → 编辑 → 导出 QCT / 导出 Comments。

**说明**：打包结果为 **文件夹版**（非单文件 exe），目的是启动更快、兼容性更好；分发时请务必保留整个文件夹内所有文件，不要只拷贝一个 exe。

## 目录结构

```
QCT_Tools/
├── README.md              # 本说明
├── requirements.txt       # 依赖
├── requirements-dev.txt   # 打包用依赖（含 pyinstaller）
├── QCT_Tools.spec         # PyInstaller 打包配置
├── build.bat              # 一键打包脚本
├── qct_app.bat            # Windows 双击启动小工具（需本机有 Python）
├── 创建桌面快捷方式.bat    # 在桌面生成可用的 qct_app.bat
├── app_gui.py             # 小工具界面：导入 PDT、编辑、导出 QCT
├── config.py              # PDT/QCT 列名、Sheet 名配置
├── pdt_reader.py          # PDT 读取（供 CLI/GUI 共用）
├── qct_data.py            # QCT 行数据构建与导出
├── qct_template.py        # QCT 空模板（表头）生成
├── create_qct_template.py  # 命令行：导出模板文件
└── pdt_to_qct.py          # 命令行：PDT → QCT 一键生成
```

## 配置修改

- **PDT 列名**与 **QCT 列名/Sheet 名** 在 `config.py` 中统一配置，若实际 PDT 或 QCT 表头与默认不一致，可修改该文件中的常量（如 `PDT_COLUMNS`、`QCT_HEADERS_SDTM`、`QCT_HEADERS_ADAM_TFL`）。

## 注意事项

- PDT 中必须包含上述 6 列，列名需与 `config.py` 中 `PDT_COLUMNS` 一致（或修改配置以匹配实际表头）。
- Output Type 为 **SDTM** 时写入第一个 Sheet，其余（如 ADaM、TFL）写入第二个 Sheet。
- 生成后的 QCT 中，QC 结果、Followup 说明等需由 QC/开发人员在 Excel 中继续填写。
