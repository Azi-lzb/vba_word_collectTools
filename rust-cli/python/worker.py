import json
import os
import sys
import traceback

for _stream in (sys.stdin, sys.stdout):
    try:
        _stream.reconfigure(encoding="utf-8", errors="strict")
    except (AttributeError, OSError):
        pass


def emit(value):
    sys.stdout.write(json.dumps(value, ensure_ascii=False))


def probe_module(name):
    try:
        __import__(name)
        return True
    except Exception:
        return False


def probe_com(progid):
    initialized = False
    app = None
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        initialized = True
        app = win32com.client.DispatchEx(progid)
        app.Visible = False
        app.DisplayAlerts = 0
        return True
    except Exception:
        return False
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        if initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


def clean(value):
    return " ".join((value or "").replace("\r", " ").replace("\n", " ").split())


def worker_probe():
    emit({
        "executable": sys.executable,
        "openpyxl": probe_module("openpyxl"),
        "pywin32": probe_module("win32com.client"),
        "word_com": probe_com("Word.Application"),
        "excel_com": probe_com("Excel.Application"),
    })


def _vba_cell(value):
    import math
    import re
    if value == "":
        return None
    try:
        number = float(value.strip())
        if math.isfinite(number):
            return number
    except (ValueError, AttributeError):
        pass
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", lambda m: "_x%04X_" % ord(m[0]), value)


def write_openpyxl(job):
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter

    report = job["report"]
    records = report["records"]
    workbook = Workbook()
    name = "output" if job["form_kind"] == "legacy" else "output_cc"
    worksheet = workbook.active
    worksheet.title = name
    bold = Font(bold=True)

    if job["layout"] == "wide" and job["form_kind"] == "legacy":
        worksheet.cell(1, 1, "源文件名 (映射后)")
        for row, source in enumerate(report["sources"], 2):
            worksheet.cell(row, 1, source["display"])
            for col, record in enumerate((r for r in records if r["source_file"] == source["file"]), 2):
                if row == 2:
                    worksheet.cell(1, col, record["column_key"])
                worksheet.cell(row, col, _vba_cell(record["answer_text"]))
        for cell in worksheet[1]:
            cell.font = bold
    elif job["layout"] == "wide":
        keys = []
        seen = set()
        for record in records:
            key = record["column_key"]
            if key not in seen:
                seen.add(key)
                keys.append(key)
        sources = [s["display"] for s in report["sources"]]
        row_index = {s["file"]:i+2 for i,s in enumerate(report["sources"])}
        values = {}
        for record in records:
            source = record["source_file"]
            if source not in row_index:
                row_index[source] = len(sources) + 2
                sources.append(record["display_source"])
            key = (source, record["column_key"])
            values[key] = record["answer_text"]
        worksheet.cell(1, 1, "源文件名")
        for column, key in enumerate(keys, 2):
            worksheet.cell(1, column, key)
        for cell in worksheet[1]:
            cell.font = bold
        for source, row in row_index.items():
            worksheet.cell(row, 1, next(s["display"] for s in report["sources"] if s["file"] == source))
            for column, key in enumerate(keys, 2):
                value = values.get((source, key))
                if value is not None:
                    worksheet.cell(row, column, _vba_cell(value))
    else:
        worksheet.append(["源文件名", "题目键", "题目", "答案"])
        for cell in worksheet[1]:
            cell.font = bold
        for record in records:
            worksheet.append([
                record["display_source"], record["column_key"],
                record["question_text"], record["answer_text"],
            ])

    if job["table_mode"] == "separate-sheet" and any(record["tables"] for record in records):
        tables = workbook.create_sheet("Tables")
        tables.append(["源文件名", "题目键", "表格序号", "行", "列", "单元格内容"])
        for cell in tables[1]:
            cell.font = bold
        for record in records:
            for table in record["tables"]:
                for row_num, cells in enumerate(table["rows"], 1):
                    for col_num, value in enumerate(cells, 1):
                        tables.append([
                            record["display_source"], record["column_key"],
                            table["index"], row_num, col_num, value,
                        ])
        for column in range(1, tables.max_column + 1):
            width = min(48, max(12, max(len(str(tables.cell(row, column).value or "")) for row in range(1, tables.max_row + 1)) + 2))
            tables.column_dimensions[get_column_letter(column)].width = width

    for sheet in workbook.worksheets:
        sheet.freeze_panes = "B2"
        for column in range(1, sheet.max_column + 1):
            width = min(60, max(14, max(len(str(sheet.cell(row, column).value or "")) for row in range(1, sheet.max_row + 1)) + 2))
            sheet.column_dimensions[get_column_letter(column)].width = width
    workbook.save(job["output"])


def _com_text(value):
    return str(value or "").replace("\x07", "").strip()


def _append_text(target_doc, text, bold=False):
    target = target_doc.Content
    target.Collapse(0)
    start = target.End
    target.InsertAfter(str(text))
    target = target_doc.Range(start, target_doc.Content.End - 1)
    if bold:
        target.Font.Bold = True
    target.Collapse(0)
    return target


def _append_range(target_doc, source_range):
    target = target_doc.Content
    target.Collapse(0)
    source_range.Copy()
    target.Paste()
    target.Collapse(0)
    return target


def _tag(cc, priority):
    if priority == "title":
        return str(cc.Title or cc.Tag or "未命名控件")
    return str(cc.Tag or cc.Title or "未命名控件")


def _unprotect(doc, password):
    protection_type = doc.ProtectionType
    if protection_type == -1:
        return protection_type, False
    try:
        doc.Unprotect(Password=password)
        return protection_type, True
    except Exception:
        return protection_type, False


def _open_word_doc(word, path):
    return word.Documents.Open(
        FileName=os.path.abspath(path), ReadOnly=False, Visible=False,
        AddToRecentFiles=False,
    )


def _read_records(word, job, opened_documents):
    from pathlib import Path
    records = []
    job["_sources"] = []
    for path in job["inputs"]:
        doc = _open_word_doc(word, path)
        state = {"doc": doc, "protection_type": -1, "unprotected": False}
        opened_documents.append(state)
        display = Path(path).name
        if job['form_kind'] == 'legacy':
            for key, value in job.get('workbook_mapping', []):
                if key.lower() in display.lower():
                    display = value
                    break
        protection_type, unprotected = _unprotect(doc, job.get("password", ""))
        state["protection_type"] = protection_type
        state["unprotected"] = unprotected
        if job["form_kind"] == "legacy" and display == Path(path).name and job.get("mapping_csv") and os.path.exists(job["mapping_csv"]):
            with open(job["mapping_csv"], "r", encoding="utf-8-sig") as mapping:
                for line in mapping:
                    key, sep, value = line.strip().partition(",")
                    if sep and key and key.lower() in display.lower():
                        display = value.strip().strip('"')
                        break
        job["_sources"].append({"file": os.path.abspath(path), "display": display})
        if job["form_kind"] == "legacy":
            previous_end = 0
            for index in range(1, doc.FormFields.Count + 1):
                ff = doc.FormFields.Item(index)
                question = doc.Range(previous_end, ff.Range.Start)
                question_text = str(question.Text)
                group_key = question_text.replace("\r", "").replace("\n", "").replace(":", "").replace("：", "").strip(" ") or str(ff.Name)
                answer_text = str(ff.Result)
                answer_range = None
                try:
                    if ff.Range.Fields.Count > 0:
                        answer_range = ff.Range.Fields.Item(1).Result
                except Exception:
                    pass
                records.append({
                    "group_key": group_key,
                    "column_key": str(ff.Name),
                    "question_text": question_text or str(ff.Name),
                    "source_file": os.path.abspath(path),
                    "display_source": display,
                    "answer_text": answer_text,
                    "_question_range": question if group_key != str(ff.Name) or question_text.replace("\r", "").strip() else None,
                    "_answer_range": answer_range,
                    "_kind": "legacy",
                })
                previous_end = ff.Range.End
        else:
            previous_end = 0
            for index in range(1, doc.ContentControls.Count + 1):
                cc = doc.ContentControls.Item(index)
                question = None
                question_text = ""
                try:
                    if cc.Range.Start >= previous_end:
                        question = doc.Range(previous_end, cc.Range.Start)
                        question_text = str(question.Text)
                except Exception:
                    question = None
                key = _tag(cc, job["priority"])
                records.append({
                    "group_key": key if cc.Tag or cc.Title else "未命名题目",
                    "column_key": key,
                    "question_text": question_text or key,
                    "source_file": os.path.abspath(path),
                    "display_source": display,
                    "answer_text": str(cc.Range.Text),
                    "_question_range": question if question_text else None,
                    "_answer_range": cc.Range,
                    "_kind": "content-control",
                })
                previous_end = max(previous_end, cc.Range.End)
    return records


def _excel_com(excel, job, records):
    from collections import OrderedDict
    from pathlib import Path

    workbook = excel.Workbooks.Add()
    sheet = workbook.Worksheets.Item(1)
    while workbook.Worksheets.Count > 1:
        workbook.Worksheets.Item(workbook.Worksheets.Count).Delete()
    sheet.Name = "output" if job["form_kind"] == "legacy" else "output_cc"
    groups = OrderedDict((s["file"], {"name":s["display"],"values":{}}) for s in job["_sources"])
    keys = []
    for record in records:
        key = record["column_key"]
        if key not in keys:
            keys.append(key)
        groups.setdefault(record["source_file"], {"name": record["display_source"], "values": {}})
        values = groups[record["source_file"]]["values"]
        values[key] = record["answer_text"]

    if job["layout"] == "wide" and job["form_kind"] == "legacy":
        sheet.Cells(1, 1).Value = "源文件名 (映射后)"
        for row, source in enumerate(job["_sources"], 2):
            sheet.Cells(row, 1).Value = source["display"]
            for col, record in enumerate((r for r in records if r["source_file"] == source["file"]), 2):
                if row == 2:
                    sheet.Cells(1, col).Value = record["column_key"]
                sheet.Cells(row, col).Value = record["answer_text"]
        sheet.Range("A1:Z1").Font.Bold = True
    elif job["layout"] == "wide":
        sheet.Cells(1, 1).Value = "源文件名"
        for col, key in enumerate(keys, 2):
            sheet.Cells(1, col).Value = key
        row = 2
        for value in groups.values():
            sheet.Cells(row, 1).Value = value["name"]
            for col, key in enumerate(keys, 2):
                if key in value["values"]:
                    sheet.Cells(row, col).Value = value["values"][key]
            row += 1
        sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, len(keys) + 1)).Font.Bold = True
    else:
        for col, title in enumerate(("源文件名", "题目键", "题目", "答案"), 1):
            sheet.Cells(1, col).Value = title
        for row, record in enumerate(records, 2):
            for col, value in enumerate((record["display_source"], record["column_key"], record["question_text"], record["answer_text"]), 1):
                sheet.Cells(row, col).Value = value
        sheet.Range("A1:D1").Font.Bold = True
    sheet.Columns.AutoFit()
    if job["table_mode"] == "separate-sheet":
        table_rows = []
        for record in records:
            rng = record.get("_answer_range")
            if record["_kind"] != "content-control" or rng is None:
                continue
            try:
                for table_index in range(1, rng.Tables.Count + 1):
                    table = rng.Tables.Item(table_index)
                    for row in range(1, table.Rows.Count + 1):
                        for col in range(1, table.Columns.Count + 1):
                            text = _com_text(table.Cell(row, col).Range.Text)
                            table_rows.append([record["display_source"], record["column_key"], table_index, row, col, text])
            except Exception:
                continue
        if table_rows:
            tables = workbook.Worksheets.Add(After=sheet)
            tables.Name = "Tables"
            headers = ["源文件名", "题目键", "表格序号", "行", "列", "单元格内容"]
            for col, value in enumerate(headers, 1):
                tables.Cells(1, col).Value = value
            for row, table_row in enumerate(table_rows, 2):
                for col, value in enumerate(table_row, 1):
                    tables.Cells(row, col).Value = value
            tables.Range("A1:F1").Font.Bold = True
            tables.Columns.AutoFit()
    output = os.path.abspath(job["output"])
    extension = Path(output).suffix.lower()
    file_format = 56 if extension == ".xls" else (52 if extension == ".xlsm" else 51)
    workbook.SaveAs(output, FileFormat=file_format)
    workbook.Close(SaveChanges=True)


def _word_com(word, job, records):
    from collections import OrderedDict
    from pathlib import Path

    document = word.Documents.Add()
    def text_at_end(doc, value):
        end = doc.Content.End - 1
        doc.Range(end, end).Text = value

    def formatted_at_end(doc, source):
        end = doc.Content.End - 1
        doc.Range(end, end).FormattedText = source.FormattedText

    if job["form_kind"] == "legacy":
        groups = OrderedDict()
        for record in records:
            groups.setdefault(record["group_key"], []).append(record)
        for group, entries in groups.items():
            section = word.Documents.Add(Visible=False)
            prompt = entries[0].get("_question_range")
            if prompt is not None:
                formatted_at_end(section, prompt)
                if not str(prompt.Text).endswith("\r"):
                    text_at_end(section, "\r")
            else:
                text_at_end(section, group + "\r")
            for record in entries:
                text_at_end(section, "【" + record["display_source"] + "】: ")
                answer = record.get("_answer_range")
                if answer is not None:
                    formatted_at_end(section, answer)
                else:
                    text_at_end(section, record["answer_text"])
                text_at_end(section, "; \r")
            formatted_at_end(document, section.Content)
            section.Close(SaveChanges=0)
    else:
        seen = set()
        for record in records:
            target = document.Content
            target.Collapse(0)
            key = record["group_key"].casefold()
            if key not in seen:
                prompt = record.get("_question_range")
                if prompt is not None and _com_text(prompt.Text):
                    prompt.Copy()
                    target.Paste()
                    target.Collapse(0)
                    if not str(prompt.Text).endswith("\r"):
                        target.InsertAfter("\r")
                    target.Collapse(0)
                else:
                    target.InsertAfter(record["group_key"] + "\r")
                    target.Font.Bold = True
                    target.Collapse(0)
                seen.add(key)
            target.InsertAfter(record["display_source"] + "\r")
            target.Font.Bold = False
            target.Collapse(0)
            record["_answer_range"].Copy()
            target.Paste()
            target.Collapse(0)
            target.InsertAfter("\r")
            target.Collapse(0)

    output = os.path.abspath(job["output"])
    extension = Path(output).suffix.lower()
    file_format = 13 if extension == ".docm" else (0 if extension == ".doc" else 12)
    document.SaveAs2(FileName=output, FileFormat=file_format)
    document.Close(SaveChanges=0)


def com_extract(job):
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = None
    excel = None
    opened_documents = []
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        records = _read_records(word, job, opened_documents)
        if job["output_kind"] == "word":
            _word_com(word, job, records)
        else:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = 0
            _excel_com(excel, job, records)
    finally:
        for state in reversed(opened_documents):
            doc = state["doc"]
            if state["protection_type"] != -1 and state["unprotected"]:
                try:
                    doc.Protect(
                        Type=state["protection_type"],
                        NoReset=True,
                        Password=job.get("password", ""),
                    )
                except Exception:
                    pass
            try:
                doc.Close(SaveChanges=0)
            except Exception:
                pass
        if excel is not None:
            excel.Quit()
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()


def main():
    operation = sys.argv[1]
    if operation == "probe":
        worker_probe()
    elif operation == "probe-openpyxl":
        emit({"available": probe_module("openpyxl")})
    elif operation == "probe-word-com":
        emit({"available": probe_com("Word.Application")})
    elif operation == "probe-excel-com":
        emit({"available": probe_com("Excel.Application")})
    elif operation == "write-excel":
        write_openpyxl(json.load(sys.stdin))
    elif operation == "com-extract":
        com_extract(json.load(sys.stdin))
    else:
        raise RuntimeError("Unknown operation: " + operation)


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        traceback.print_exc(file=sys.stderr)
        sys.exit(1)
