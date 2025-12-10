using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeMcpServer.Tools;

public class WordAddIndexTool : IAsposeTool
{
    public string Description => "Add index (XE fields and INDEX field) to Word document";

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
            indexEntries = new
            {
                type = "array",
                description = "Array of index entries (objects with 'text' and optional 'subEntry', 'pageRangeBookmark')",
                items = new
                {
                    type = "object",
                    properties = new
                    {
                        text = new { type = "string" },
                        subEntry = new { type = "string" },
                        pageRangeBookmark = new { type = "string" }
                    },
                    required = new[] { "text" }
                }
            },
            insertIndexAtEnd = new
            {
                type = "boolean",
                description = "Insert INDEX field at end of document (optional, default: true)"
            },
            headingStyle = new
            {
                type = "string",
                description = "Heading style for index (optional, default: 'Heading 1')"
            }
        },
        required = new[] { "path", "indexEntries" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var indexEntriesArray = arguments?["indexEntries"]?.AsArray() ?? throw new ArgumentException("indexEntries is required");
        var insertIndexAtEnd = arguments?["insertIndexAtEnd"]?.GetValue<bool?>() ?? true;
        var headingStyle = arguments?["headingStyle"]?.GetValue<string>() ?? "Heading 1";

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);

        // Insert XE fields for index entries
        foreach (var entryObj in indexEntriesArray)
        {
            if (entryObj is JsonObject entry)
            {
                var text = entry["text"]?.GetValue<string>();
                var subEntry = entry["subEntry"]?.GetValue<string>();
                var pageRangeBookmark = entry["pageRangeBookmark"]?.GetValue<string>();

                if (!string.IsNullOrEmpty(text))
                {
                    builder.MoveToDocumentEnd();
                    var xeField = $"XE \"{text}\"";
                    if (!string.IsNullOrEmpty(subEntry))
                    {
                        xeField += $" \\t \"{subEntry}\"";
                    }
                    if (!string.IsNullOrEmpty(pageRangeBookmark))
                    {
                        xeField += $" \\r \"{pageRangeBookmark}\"";
                    }
                    builder.InsertField(xeField);
                }
            }
        }

        // Insert INDEX field at end
        if (insertIndexAtEnd)
        {
            builder.MoveToDocumentEnd();
            builder.InsertBreak(BreakType.PageBreak);
            builder.ParagraphFormat.Style = doc.Styles[headingStyle];
            builder.Writeln("Index");
            builder.ParagraphFormat.Style = doc.Styles["Normal"];
            builder.InsertField("INDEX \\e \" \" \\h \"A\"");
        }

        doc.Save(path);
        return await Task.FromResult($"Index entries added. Total entries: {indexEntriesArray.Count}");
    }
}

