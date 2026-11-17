# Aspose.NET 实验及耗时记录

Windows x64；Aspose.Words / Aspose.Cells 26.10.0，.NET 8，未提供许可证。每次均新建 CLI 进程；Aspose 路由再启动 .NET 子进程。每组先预热一次，3 份文档记录 5 次、30 份记录 3 次，以下为总耗时中位数。30 份由 3 份生成示例各复制 10 次，属于小型合成样本。

Excel 测试主数据表单元格一致，但 Aspose 试用结果多出 Evaluation Warning 工作表；Word 的 Aspose 原型仅输出文本、保留的内容及格式与 Rust 不等价，速度仅供参考。

以下每行依次为 Rust 与 Aspose 耗时：

- 3 份 · legacy → excel：15.3 ms / 1357.5 ms。
- 3 份 · legacy → word：21.4 ms / 1691.4 ms。
- 3 份 · content-control → excel：18.4 ms / 1741.7 ms。
- 3 份 · content-control → word：22.9 ms / 1781.2 ms。
- 30 份 · legacy → excel：25.2 ms / 2369.9 ms。
- 30 份 · legacy → word：49.9 ms / 2696.7 ms。
- 30 份 · content-control → excel：26.9 ms / 2261.3 ms。
- 30 份 · content-control → word：46.3 ms / 2512.2 ms。

Aspose 实验后端不支持源 Word 格式、表格、图片、超链接保留，暂不支持 TITLE 优先级和文件名映射。试用版本存在文档大小限制，不能由这些数字推断大文件、正式授权版或常驻进程的性能。

[3 份完整记录](benchmark-report.json) · [30 份完整记录](benchmark-report-30-files.json) · [构建与复现说明](README.md)
