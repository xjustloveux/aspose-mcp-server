using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeMcpServer.Tools;

public class WordInsertFieldTool : IAsposeTool
{
    public string Description => "Insert a field (DATE, TIME, PAGE, NUMPAGES, AUTHOR, FILENAME, etc.) into a Word document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Document file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input)"
            },
            fieldType = new
            {
                type = "string",
                description = "Field type: DATE, TIME, PAGE, NUMPAGES, AUTHOR, FILENAME, TITLE, SUBJECT, COMPANY, CREATEDATE, SAVEDATE, PRINTDATE, MERGEFIELD, etc.",
                @enum = new[] { "DATE", "TIME", "PAGE", "NUMPAGES", "AUTHOR", "FILENAME", "TITLE", "SUBJECT", 
                               "COMPANY", "CREATEDATE", "SAVEDATE", "PRINTDATE", "USERNAME", "DOCPROPERTY", 
                               "HYPERLINK", "REF", "MERGEFIELD", "IF", "SEQ" }
            },
            fieldArgument = new
            {
                type = "string",
                description = "Field argument (e.g., date format '\\@ \"yyyy-MM-dd\"', property name for DOCPROPERTY, bookmark name for REF)"
            },
            paragraphIndex = new
            {
                type = "number",
                description = "Paragraph index to insert into (0-based). If not provided, inserts at end of document. Use -1 to insert at beginning."
            },
            insertAtStart = new
            {
                type = "boolean",
                description = "Insert at the start of the paragraph (default: false, inserts at end)"
            }
        },
        required = new[] { "path", "fieldType" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var fieldType = arguments?["fieldType"]?.GetValue<string>()?.ToUpper() ?? throw new ArgumentException("fieldType is required");
        var fieldArgument = arguments?["fieldArgument"]?.GetValue<string>() ?? "";
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int?>();
        var insertAtStart = arguments?["insertAtStart"]?.GetValue<bool>() ?? false;

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);

        // Determine insertion position
        if (paragraphIndex.HasValue)
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            if (paragraphIndex.Value == -1)
            {
                if (paragraphs.Count > 0)
                {
                    var firstPara = paragraphs[0] as Paragraph;
                    if (firstPara != null)
                    {
                        builder.MoveTo(firstPara);
                        if (insertAtStart)
                            builder.MoveToBookmark(firstPara.Range.Bookmarks.Count > 0 ? firstPara.Range.Bookmarks[0].Name : "");
                        else
                            builder.MoveToParagraph(0, firstPara.GetText().Length);
                    }
                }
            }
            else if (paragraphIndex.Value >= 0 && paragraphIndex.Value < paragraphs.Count)
            {
                var targetPara = paragraphs[paragraphIndex.Value] as Paragraph;
                if (targetPara != null)
                {
                    builder.MoveTo(targetPara);
                    if (!insertAtStart)
                    {
                        builder.MoveToParagraph(paragraphIndex.Value, targetPara.GetText().Length);
                    }
                }
                else
                {
                    throw new InvalidOperationException($"無法找到索引 {paragraphIndex.Value} 的段落");
                }
            }
            else
            {
                throw new ArgumentException($"段落索引 {paragraphIndex.Value} 超出範圍 (文檔共有 {paragraphs.Count} 個段落)");
            }
        }
        else
        {
            builder.MoveToDocumentEnd();
        }

        // Build field code
        string fieldCode = fieldType;
        if (!string.IsNullOrEmpty(fieldArgument))
        {
            fieldCode += " " + fieldArgument;
        }

        // Insert field
        Field field;
        try
        {
            field = builder.InsertField(fieldCode);
            
            // Update the field to show its result
            field.Update();
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"無法插入欄位 '{fieldCode}': {ex.Message}", ex);
        }

        doc.Save(outputPath);

        var result = $"成功插入欄位\n";
        result += $"欄位類型: {fieldType}\n";
        if (!string.IsNullOrEmpty(fieldArgument))
            result += $"欄位參數: {fieldArgument}\n";
        result += $"欄位代碼: {fieldCode}\n";
        
        try
        {
            var fieldResult = field.Result;
            if (!string.IsNullOrEmpty(fieldResult))
                result += $"欄位結果: {fieldResult}\n";
        }
        catch { }
        
        if (paragraphIndex.HasValue)
        {
            if (paragraphIndex.Value == -1)
                result += "插入位置: 文檔開頭\n";
            else
                result += $"插入位置: 段落 #{paragraphIndex.Value}\n";
        }
        else
        {
            result += "插入位置: 文檔末尾\n";
        }
        
        result += $"輸出: {outputPath}";

        return await Task.FromResult(result);
    }
}

