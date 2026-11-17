# Word汇总工具

批量提取 Word 问卷中的**旧式窗体域**和**内容控件**，汇总到 Excel 或 Word。提供可直接运行的 Rust 程序，也保留原有 Excel/Word VBA 实现。

## 下载

**Windows 用户推荐：[下载 Windows ZIP（v0.1.0，x64）](https://github.com/Azi-lzb/vba_word_collectTools/releases/download/v0.1.0/WordSummaryTool-0.1.0-windows-x86_64.zip)**，解压后即可使用。

- [最新版本下载入口](https://github.com/Azi-lzb/vba_word_collectTools/releases/latest)
- [Windows 独立 EXE](https://github.com/Azi-lzb/vba_word_collectTools/releases/download/v0.1.0/WordSummaryTool.exe) · [空密码配置 config.xlsx](https://github.com/Azi-lzb/vba_word_collectTools/releases/download/v0.1.0/config.xlsx)
- [Linux tar.gz 便携包](https://github.com/Azi-lzb/vba_word_collectTools/releases/download/v0.1.0/WordSummaryTool-0.1.0-linux-x86_64.tar.gz)
- [Debian / Ubuntu DEB 安装包](https://github.com/Azi-lzb/vba_word_collectTools/releases/download/v0.1.0/WordSummaryTool_0.1.0_amd64.deb)
- [SHA256 校验文件](https://github.com/Azi-lzb/vba_word_collectTools/releases/download/v0.1.0/SHA256SUMS.txt)

GitHub 附件名使用 `WordSummaryTool`；ZIP 内的程序名为 `Word汇总工具.exe`。Windows ZIP 只包含程序和 `config.xlsx` 两个文件。Linux 包适用于 **x86_64、glibc 2.28 及以上**，不适用于 ARM64 或 Alpine/musl。

## Windows 快速使用

1. 下载上面的 Windows ZIP，将两个文件解压到同一文件夹。
2. 双击 `Word汇总工具.exe`，进入终端交互菜单。
3. 选择 **1 文件提取**，再按问卷类型选择“旧式窗体”或“内容控件”，以及输出到 Excel 或 Word。
4. 选择多个文件、文件夹当前层或递归子目录；也可以手动输入路径。选完后程序自动开始提取。
5. 结果保存在程序旁边的 `output` 文件夹。同名文件会自动编号，不覆盖已有结果。

主菜单分为三组：

- **1 文件提取**：旧式窗体 → Excel / Word、内容控件 → Excel / Word。
- **2 示例文档生成**：生成三份旧式窗体示例、三份内容控件示例，或两种一起生成。示例默认放在 VBA 工作簿所在目录；没有工作簿时放在程序目录。
- **3 常用操作**：打开输出目录、打开最新输出文件、打开配置。子菜单输入 `0` 可返回。

### 配置密码与文件名映射

进入 **3 常用操作 → 3 打开配置**，编辑并保存 `config.xlsx`，下次提取时生效：

- `config!B2`：Word 文档的“限制编辑”密码，没有密码时留空。**Release 中的密码已清空，请填写自己的密码。**
- `config!B3`：内容控件优先按 `TAG` 或 `TITLE` 识别，默认 `TAG`。
- `mapping!A:B`：从第 2 行起填写“文件名关键词”和“映射名称”，用于旧式窗体汇总的来源名称。

独立配置可直接使用，无需准备原 VBA 工作簿。

### 运行环境与格式

日常处理 `.docx` / `.docm`、输出 `.xlsx` / `.docx` 时，程序优先使用**纯 Rust**，不需要 Python 或 Office。Word 汇总支持已覆盖的文字格式、图片、超链接和表格；复杂文档建议检查汇总结果。

自动路由无需手动选择：Excel 写出失败时依次尝试 openpyxl、COM；Word 失败时尝试 COM。Rust 读取失败时直接回退 COM，因为 openpyxl 共用 Rust 读取器。旧二进制 `.doc` 和旧式输出格式需要 Windows COM。

可选回退环境：openpyxl 需要 Python 和 `openpyxl`；COM 需要 Python、`pywin32` 和 Word，输出 Excel 时还需要 Excel。**旧 `.doc` 在未安装 Office 的环境中需先转换为 `.docx`。**

## Linux 快速使用

解压便携包后，在终端运行：

```bash
tar -xzf WordSummaryTool-0.1.0-linux-x86_64.tar.gz
cd Word汇总工具-0.1.0-linux-x86_64
./Word汇总工具
```

Debian / Ubuntu 也可安装 DEB：

```bash
sudo apt install ./WordSummaryTool_0.1.0_amd64.deb
Word汇总工具
```

Linux 交互菜单使用手动输入路径。便携版结果写入程序旁的 `output`；安装到 `/usr/bin` 的版本写入当前目录的 `output`。COM 仅适用于 Windows。

更多命令行参数、构建和验证方法见 [Rust 工具说明](rust-cli/README.md) 与 [打包说明](rust-cli/packaging/README.md)。

---

## 原 VBA 版本文档

以下说明适用于 Excel/Word VBA 宏版本。VBA 套件主要用于受保护（限制编辑：填写窗体）的 Word 问卷，可汇总到 Excel 表格或 Word 报告。

## 🌟 核心功能

### 1. 旧式窗体汇总 (Legacy Form Fields)
- **`extractWord.bas`**: 将旧式窗体域内容汇总至 Excel `output` 表。支持文件名映射。
- **`extractToWord.bas`**: 将旧式窗体内容按题目汇总至新 Word 文档。自动识别窗体前的文字作为题目，并保留题目和答案在源文档中的字体、字号、颜色等格式。

### 2. 新式窗体/内容控件汇总 (Content Controls)
- **`extractCCtoExcel.bas`**: 按 `Tag` 或 `Title` 将内容控件数据汇总至 Excel。
- **`extractCCtoWord.bas`**: 汇总内容控件至 Word（文本模式）。
- **`extractCCtoWordWithTables.bas`**: **(推荐)** 汇总内容控件至 Word（含表格模式）。按 `Tag/Title` 汇总，复制源文档中控件前的题目文本及其格式，并保留控件内表格、行列和格式；若找不到题目文本，则以 `Tag/Title` 作为标题。

### 3. 配置与工具
- **`modConfig.bas`**: **(新)** 提供 `config` 工作表初始化及配置读取功能。
- **`GenerateTestData.bas` / `GenerateCCTestData.bas`**: 生成模拟问卷样本。

---

## 📘 核心指南：旧式窗体 vs 内容控件

在设计问卷和选择汇总工具前，请务必阅读以下差异总结：

### 1. 定位逻辑与底层技术
- **旧式窗体 (Legacy)**：**依赖“字典 (Dictionary) 字符串匹配”**。代码读取窗体前的文字存入内存字典。
    - *风险*：题目中哪怕多一个空格或标点不统一，都会被识别为两个不同的题目。
- **内容控件 (New)**：**依赖“对象属性 (Tag/Title) 定位”**。
    - *优势*：**精准且稳定**。无论文档表面的题目文字如何变动，只要后台标识一致，数据就能完美对齐。

### 2. 表格支持能力
- **旧式窗体**：无法在填空位内嵌套表格。只能提取纯文本答案。
- **内容控件**：支持表格嵌套。推荐使用 `extractCCtoWordWithTables` 模块进行完整复制。

---

## ⚙️ 全新配置系统 (config 工作表)

本项目现已启用独立的配置管理系统。您可以通过以下步骤进行配置：

1.  **初始化**：点击 Excel 菜单中的 **“vbaWord工具” -> “3. 配置与工具” -> “3.1 初始化 config 配置表”**。
2.  **填写参数**：系统会自动生成 `config` 工作表，请在 **B 列** 填写对应的值：

| 配置项 (A列) | 功能说明 (C列) | 可填值示例 (B列) |
| :--- | :--- | :--- |
| **文档保护密码** | Word 文档的“限制编辑”密码 | `123456` (若无请留空) |
| **题目识别优先级** | 内容控件的提取逻辑 | `TAG` (标记优先) 或 `TITLE` (标题优先) |

- **优先级逻辑**：若设为 `TITLE`，程序优先提取“标题”；若设为 `TAG` 或留空，则优先提取“标记”。若首选属性为空，程序会自动回退使用备选属性。

---

## ⚠️ 避坑指南与最佳实践

### 1. 内容控件的“唯一性”
- **警告**：直接复制粘贴内容控件会连同 `Tag` 一起复制。请务必在粘贴后修改其 `Tag`，确保在一个模板内没有重复的 `Tag`，否则数据会被合并。

### 2. 密码支持
- 只有在密码正确且成功解除保护的情况下，代码才会在处理完后自动恢复保护，彻底避免 **3012 错误**。

### 3. 导入与编码
- 本项目所有 `.bas` 和 `.cls` 文件均使用 **ANSI (GBK)** 编码。请直接使用 VBA 编辑器的“导入文件”功能，切勿将其转存为 UTF-8。

---

## 🛠️ 如何使用

1.  **准备工作**：在 VBA 编辑器中勾选 `Microsoft Word 16.0 Object Library` 和 `Microsoft Scripting Runtime`。
2.  **配置**：运行“初始化 config 配置表”并填写参数。
3.  **运行**：根据问卷类型选择对应的汇总宏，批量选择文件即可。

---

## 📂 项目结构
- `VBA_Export/Modules/`: 核心 VBA 模块。
- `VBA_Export/Documents/`: `ThisWorkbook` 菜单管理。
- `SampleForTest/`: 测试样本。

## ✒️ 开发者
- **Azi-lzb**
