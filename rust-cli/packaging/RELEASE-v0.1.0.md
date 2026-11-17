Word汇总工具 v0.1.0

- Windows x64：下载 ZIP 解压后运行 `Word汇总工具.exe`，或单独下载 EXE 和 `config.xlsx` 放在同一目录。
- ZIP 仅包含 EXE 和配置工作簿。发布配置的文档保护密码已清空；请在“常用操作 → 打开配置”中填写自己的设置并保存。
- Linux x86_64：可下载 tar.gz 便携包，或在 Debian/Ubuntu 安装 DEB。要求 glibc 2.28 及以上，不适用于 ARM64 或 Alpine/musl。

支持旧式窗体和内容控件提取到 Excel、Word，以及示例文档生成。选完文件后自动执行，结果保存在 `output`，同名文件自动编号。菜单中可打开输出目录、最新结果和配置。

自动优先纯 Rust；Excel 写出失败时尝试 openpyxl，再尝试 COM；Word 失败时回退 COM。旧 `.doc` 等格式需要 Windows 上的 Word/Excel COM。纯 Rust 不需要 Python 或 Office；可选回退需要自行安装相应组件。Aspose 已从程序移除。

EXE、ZIP、tar.gz、DEB、空密码配置和 SHA256 校验文件均可在下方下载。Windows 和 Ubuntu 20.04/22.04 的提取、安装及卸载已验证。
