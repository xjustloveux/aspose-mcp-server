using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeMcpServer.Tools;

public class WordAddTableOfContentsTool : IAsposeTool
{
    public string Description => "Add a table of contents to a Word document";

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
            position = new
            {
                type = "string",
                description = "Insert position: start, end (default: start)",
                @enum = new[] { "start", "end" }
            },
            title = new
            {
                type = "string",
                description = "Table of contents title (default: '目錄')"
            },
            maxLevel = new
            {
                type = "number",
                description = "Maximum heading level to include (1-9, default: 3)"
            },
            hyperlinks = new
            {
                type = "boolean",
                description = "Enable clickable hyperlinks (default: true)"
            },
            pageNumbers = new
            {
                type = "boolean",
                description = "Show page numbers (default: true)"
            },
            rightAlignPageNumbers = new
            {
                type = "boolean",
                description = "Right-align page numbers (default: true)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var position = arguments?["position"]?.GetValue<string>() ?? "start";
        var title = arguments?["title"]?.GetValue<string>() ?? "目錄";
        var maxLevel = arguments?["maxLevel"]?.GetValue<int>() ?? 3;
        var hyperlinks = arguments?["hyperlinks"]?.GetValue<bool>() ?? true;
        var pageNumbers = arguments?["pageNumbers"]?.GetValue<bool>() ?? true;
        var rightAlignPageNumbers = arguments?["rightAlignPageNumbers"]?.GetValue<bool>() ?? true;

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);

        // Move to position
        if (position == "end")
        {
            builder.MoveToDocumentEnd();
        }
        else
        {
            builder.MoveToDocumentStart();
        }

        // Add title
        if (!string.IsNullOrEmpty(title))
        {
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln(title);
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        }

        // Build TOC field code
        var switches = new List<string>();
        switches.Add($"\\o \"1-{maxLevel}\""); // Outline levels

        if (!hyperlinks)
            switches.Add("\\n"); // No hyperlinks

        if (!pageNumbers)
            switches.Add("\\n"); // No page numbers
        else if (rightAlignPageNumbers)
            switches.Add("\\p \" \""); // Right align page numbers

        var tocFieldCode = $"TOC {string.Join(" ", switches)}";

        // Insert TOC field
        var field = builder.InsertField(tocFieldCode, "");

        // Update field
        field.Update();

        doc.Save(outputPath);

        var result = $"成功添加目錄\n";
        if (!string.IsNullOrEmpty(title)) result += $"標題: {title}\n";
        result += $"位置: {(position == "start" ? "文檔開頭" : "文檔結尾")}\n";
        result += $"最大層級: {maxLevel}\n";
        result += $"超連結: {(hyperlinks ? "是" : "否")}\n";
        result += $"頁碼: {(pageNumbers ? "是" : "否")}\n";
        result += $"輸出: {outputPath}\n";
        result += "\n注意: 在 Word 中打開文檔後，可能需要右鍵點擊目錄選擇「更新域」以更新頁碼";

        return await Task.FromResult(result);
    }
}

