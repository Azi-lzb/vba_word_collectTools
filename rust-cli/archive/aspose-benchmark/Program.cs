// Experimental baseline. Word output is text-only and does not preserve source formatting.
using System.Text;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Markup;
using Cells = Aspose.Cells;

Console.OutputEncoding = Encoding.UTF8;
try
{
    if (args.Length < 4 || !new[] { "legacy", "content-control" }.Contains(args[0]) || !new[] { "excel", "word" }.Contains(args[1]))
        throw new ArgumentException("Usage: aspose-benchmark <legacy|content-control> <excel|word> <output> <input>...");
    var wordsLicense = Environment.GetEnvironmentVariable("ASPOSE_WORDS_LICENSE");
    var cellsLicense = Environment.GetEnvironmentVariable("ASPOSE_CELLS_LICENSE");
    if (!string.IsNullOrEmpty(wordsLicense)) new Aspose.Words.License().SetLicense(wordsLicense);
    if (!string.IsNullOrEmpty(cellsLicense)) new Cells.License().SetLicense(cellsLicense);
    string kind=args[0], outputKind=args[1], output=Path.GetFullPath(args[2]);
    var files=args.Skip(3).Select(Path.GetFullPath).ToArray();
    if (files.Any(f=>string.Equals(f, output, StringComparison.OrdinalIgnoreCase))) throw new ArgumentException("Output would overwrite input");
    var documents = new List<(string File, List<Entry> Records)>();
    foreach (string file in files)
    {
        var document = new Document(file);
        var entries = new List<Entry>();
        if (kind == "legacy")
        {
            foreach (FormField field in document.Range.FormFields)
            {
                string question = field.ParentParagraph?.Range.Text ?? field.Name;
                question = question.Split('\u0013')[0].Trim('\r','\n','\a',' ',':','：');
                entries.Add(new Entry(field.Name, question, field.Result ?? ""));
            }
        }
        else
        {
            foreach (StructuredDocumentTag tag in document.GetChildNodes(NodeType.StructuredDocumentTag, true))
            {
                string key=!string.IsNullOrEmpty(tag.Tag)?tag.Tag:!string.IsNullOrEmpty(tag.Title)?tag.Title:"未命名控件";
                string text=tag.Range.Text.TrimEnd('\r','\n','\a');
                entries.Add(new Entry(key, key, text));
            }
        }
        documents.Add((Path.GetFileName(file), entries));
    }
    if (outputKind == "excel")
    {
        var book=new Cells.Workbook();
        var sheet=book.Worksheets[0];sheet.Name=kind=="legacy"?"output":"output_cc";
        sheet.Cells[0,0].PutValue(kind=="legacy"?"源文件名 (映射后)":"源文件名");
        var columns=new List<string>();
        if (kind=="legacy") columns.AddRange(documents.FirstOrDefault().Records?.Select(e=>e.Key) ?? Enumerable.Empty<string>());
        else foreach (var doc in documents) foreach(var item in doc.Records) if(!columns.Contains(item.Key)) columns.Add(item.Key);
        for(int i=0;i<columns.Count;i++) sheet.Cells[0,i+1].PutValue(columns[i]);
        for(int row=0;row<documents.Count;row++)
        {
            var doc=documents[row];sheet.Cells[row+1,0].PutValue(doc.File);
            for(int i=0;i<doc.Records.Count;i++)
            {
                var entry=doc.Records[i];int col=kind=="legacy"?i:columns.IndexOf(entry.Key);
                if(col>=0) sheet.Cells[row+1,col+1].PutValue(entry.Value);
            }
        }
        book.Save(output,Cells.SaveFormat.Xlsx);
    }
    else
    {
        var result=new Document();var builder=new DocumentBuilder(result);
        var groups=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach(var doc in documents) foreach(var entry in doc.Records)
        {
            string group=string.IsNullOrEmpty(entry.Question)?entry.Key:entry.Question;
            if(groups.Add(group)) builder.Writeln(group);
            builder.Writeln(doc.File);builder.Writeln(entry.Value);
        }
        result.Save(output,SaveFormat.Docx);
    }
    Console.WriteLine($"Aspose trial comparison: {kind}/{outputKind}, {documents.Count} documents, {documents.Sum(d=>d.Records.Count)} records");
    return 0;
}
catch(Exception error) { Console.Error.WriteLine(error); return 1; }

record Entry(string Key,string Question,string Value);
