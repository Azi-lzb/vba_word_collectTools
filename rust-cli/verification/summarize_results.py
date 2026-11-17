"""Publish an evidence-derived report; refuses incomplete or failing verification."""
import hashlib
import json
from pathlib import Path
from inspect_baseline import ROOT, OUT, snapshot

HERE=Path(__file__).resolve().parent


def read(path):
    return json.loads(path.read_text(encoding='utf-8'))


def main():
    summary=read(OUT/'verification-summary.json')
    assert summary['passed'] and summary['originals_unchanged'] and summary['repository_samples_unchanged']
    assert snapshot()==summary['original_sha256']
    for name,digest in {**summary['sample_sha256'],**summary['fixture_hashes']}.items():
        assert hashlib.sha256((ROOT/name).read_bytes()).hexdigest()==digest,name
    cases=[case for path in summary['parity_reports'] for case in read(HERE.parent/path)]
    assert all(case['passed'] for case in cases)
    normalized=[case for case in cases if case.get('raw_equal') is False]
    exact=len(cases)-len(normalized)
    graph=read(OUT/'opc-final-report.json');assert all(not item['errors'] for item in graph)
    effective=read(OUT/'rendered'/'fidelity-checks.json')
    structures=read(OUT/'fidelity-structures.json')
    assert all(check['passed'] for check in effective+structures)
    debug=read(OUT/'native-independence.json');release=read(OUT/'native-independence-release.json')
    assert debug['passed'] and release['passed']
    exe=Path(release['executable'])
    assert hashlib.sha256(exe.read_bytes()).hexdigest()==release['executable_sha256']
    failures=summary.get('expected_vba_failures',[])
    report={
        'passed':True,'completed_utc':summary['finished_utc'],
        'exact_parity_cases':exact,'normalized_rich_cases':len(normalized),
        'normalized_cases':[{'case':case['case'],'scope':case['comparison_scope'],'checks':case['checks']} for case in normalized],
        'opc_documents_checked':len(graph),'effective_and_render_checks':len(effective+structures),
        'debug_no_python_operations':len(debug['cases']),'release_no_python_operations':len(release['cases']),
        'original_files_unchanged':len(summary['original_sha256']),
        'sample_files_unchanged':len(summary['sample_sha256']),
        'vba_baseline_failures':failures,
        'release_executable':str(exe),'release_sha256':release['executable_sha256']}
    (OUT/'delivery-audit.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    failure_lines='\n'.join(f"- `{item['case']}`：{item['error']}。未计入通过数。" for item in failures) or '- 本轮没有 VBA 基线运行错误。'
    normalized_lines='\n'.join(f"- `{case['case']}`：原始输出并非逐字一致。显式归一化剪贴板产生的空段落、行末空格与图片文本标记；链接变化仅允许缺失本地书签回源及相对链接按源目录重新定位。正文、表格、图片数量以及允许的链接变化全部通过，原始 expected/actual 和各项检查均保留。" for case in normalized) or '- 本轮复杂控件 VBA 未产生可比较的完整输出，保真由源文档独立核验。'
    text=f'''# Rust CLI 验收结果

完成时间：{summary['finished_utc']}。本报告由 `summarize_results.py` 根据实际证据生成，所有必需检查通过后才会写出。

## VBA 四种输出对齐

在工作簿副本上运行真实 VBA，再用 Word/Excel 读取 CLI 实际输出；只替换文件选择、消息框和专用 Word 实例入口。没有重写 VBA 提取算法。

- 严格输出内容一致：**{exact} 项**（Rust、COM 及 Python Excel 路径；原有样本、保护/空值/复选框/下拉框/重复键/公式等边界案例）。
- 复杂控件有明确归一化范围的比较：**{len(normalized)} 项**，单独计数，不冒充逐字一致。
{normalized_lines}

VBA 本身的运行失败单独保留：

{failure_lines}

## 纯 Rust 保真

- 同名资源分源隔离、递归关系复制、图片字节一致、ID 和书签唯一性：{len(graph)} 份输出包检查通过。
- 有效字体、字号、样式继承、主题颜色、编号，以及 Word/PDF 结构与渲染：{len(effective+structures)} 项检查通过。
- 6 个实际渲染图片位置覆盖浮动、VML 与裁剪；完整核验表格网格、横纵合并、嵌套、宽度、行高、重复表头、边框与底色。
- 普通超链接、内部书签、简单/分段域、控件外目标回源、相对文件链接与外部文档书签均已核验。相对链接按源目录或 HyperlinkBase 定位，外部书签不改名。
- Word 主题明暗转换允许 RGB 各通道最多 1/255 舍入差异。未承诺任意 Word 特性或不同字体/Office 版本下的分页完全相同。

## 独立运行与原文件

- Debug 与 Release 各运行 {len(debug['cases'])} / {len(release['cases'])} 项 Rust/自动回退操作，并通过 probe。测试时 PATH 为空、配置的 Python 不存在。
- 原 VBA 导出及工作簿 **{len(summary['original_sha256'])} 个文件**、仓库样本 **{len(summary['sample_sha256'])} 个文件** SHA256 未变，测试源文档也未变。
- 没有修改现有 VBA 或 Excel；没有执行 Git 操作。
- Release：`{exe}`，SHA256 `{release['executable_sha256']}`。
- 纯 Rust 输入为 `.docx/.docm`，输出 `.docx/.xlsx`。旧二进制 `.doc` 通过 COM；Python/openpyxl 为可选 Excel 后端。

## 证据与复现

- [总执行记录](evidence/verification-summary.json)、[交付审核](evidence/delivery-audit.json)、[日志](evidence/verification-run.log)
- [资源包检查](evidence/opc-final-report.json)、[Word 有效属性与渲染检查](evidence/rendered/fidelity-checks.json)、[复杂表格及图片检查](evidence/fidelity-structures.json)
- [复杂内容控件 Rust 输出](evidence/rich-cc-rust.docx)、[Word 导出 PDF](evidence/rendered/rich-cc-rust-evidence.pdf)
- 各 `evidence/<case>/word-1.docx` 为真实 VBA Word 输出；Excel 的单元格结果在各目录 `result.json` 中。`parity-*.json` 保留完整比较。
- 复现命令和依赖见 [验证说明](README.md)，CLI 用法见 [项目说明](../README.md)。
'''
    (HERE/'RESULTS.md').write_text(text,encoding='utf-8')
    print(json.dumps(report,ensure_ascii=True))


if __name__=='__main__':main()
