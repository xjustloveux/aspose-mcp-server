using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

public class WordReferenceTool : IAsposeTool
{
    public string Description => "Manage references in Word documents (table of contents, index, cross-reference)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation: add_table_of_contents, update_table_of_contents, add_index, add_cross_reference",
                @enum = new[] { "add_table_of_contents", "update_table_of_contents", "add_index", "add_cross_reference" }
            },
            path = new
            {
                type = "string",
                description = "Document file path"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to overwrite input)"
            },
            position = new
            {
                type = "string",
                description = "Insert position: start, end (for add_table_of_contents, default: start)",
                @enum = new[] { "start", "end" }
            },
            title = new
            {
                type = "string",
                description = "Table of contents title (for add_table_of_contents, default: '目錄')"
            },
            maxLevel = new
            {
                type = "number",
                description = "Maximum heading level to include (for add_table_of_contents, default: 3)"
            },
            hyperlinks = new
            {
                type = "boolean",
                description = "Enable clickable hyperlinks (for add_table_of_contents, default: true)"
            },
            pageNumbers = new
            {
                type = "boolean",
                description = "Show page numbers (for add_table_of_contents, default: true)"
            },
            rightAlignPageNumbers = new
            {
                type = "boolean",
                description = "Right-align page numbers (for add_table_of_contents, default: true)"
            },
            tocIndex = new
            {
                type = "number",
                description = "TOC field index (0-based, for update_table_of_contents, optional)"
            },
            indexEntries = new
            {
                type = "array",
                description = "Array of index entries (for add_index)",
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
                description = "Insert INDEX field at end of document (for add_index, default: true)"
            },
            headingStyle = new
            {
                type = "string",
                description = "Heading style for index (for add_index, default: 'Heading 1')"
            },
            referenceType = new
            {
                type = "string",
                description = "Reference type: Heading, Bookmark, Figure, Table, Equation (for add_cross_reference)",
                @enum = new[] { "Heading", "Bookmark", "Figure", "Table", "Equation" }
            },
            referenceText = new
            {
                type = "string",
                description = "Text to insert before reference (for add_cross_reference, optional)"
            },
            targetName = new
            {
                type = "string",
                description = "Target name (heading text, bookmark name, etc.) (for add_cross_reference)"
            },
            insertAsHyperlink = new
            {
                type = "boolean",
                description = "Insert as hyperlink (for add_cross_reference, default: true)"
            },
            includeAboveBelow = new
            {
                type = "boolean",
                description = "Include 'above' or 'below' text (for add_cross_reference, default: false)"
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");

        return operation.ToLower() switch
        {
            "add_table_of_contents" => await AddTableOfContents(arguments),
            "update_table_of_contents" => await UpdateTableOfContents(arguments),
            "add_index" => await AddIndex(arguments),
            "add_cross_reference" => await AddCrossReference(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> AddTableOfContents(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        SecurityHelper.ValidateFilePath(path, "path");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        var position = arguments?["position"]?.GetValue<string>() ?? "start";
        var title = arguments?["title"]?.GetValue<string>() ?? "目錄";
        var maxLevel = arguments?["maxLevel"]?.GetValue<int>() ?? 3;
        var hyperlinks = arguments?["hyperlinks"]?.GetValue<bool>() ?? true;
        var pageNumbers = arguments?["pageNumbers"]?.GetValue<bool>() ?? true;
        var rightAlignPageNumbers = arguments?["rightAlignPageNumbers"]?.GetValue<bool>() ?? true;

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);

        if (position == "end")
            builder.MoveToDocumentEnd();
        else
            builder.MoveToDocumentStart();

        if (!string.IsNullOrEmpty(title))
        {
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln(title);
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        }

        var switches = new List<string>();
        switches.Add($"\\o \"1-{maxLevel}\"");

        if (!hyperlinks)
            switches.Add("\\n");

        if (!pageNumbers)
            switches.Add("\\n");

        if (!rightAlignPageNumbers)
            switches.Add("\\l");

        var fieldCode = $"TOC {string.Join(" ", switches)}";
        builder.InsertField(fieldCode);

        doc.Save(outputPath);
        return await Task.FromResult($"Table of contents added: {outputPath}");
    }

    private async Task<string> UpdateTableOfContents(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        SecurityHelper.ValidateFilePath(path, "path");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        var tocIndex = arguments?["tocIndex"]?.GetValue<int?>();

        var doc = new Document(path);
        var tocFields = doc.Range.Fields
            .Cast<Field>()
            .Where(f => f.Type == FieldType.FieldTOC)
            .ToList();

        if (tocFields.Count == 0)
            return await Task.FromResult("No table of contents fields found in document");

        if (tocIndex.HasValue)
        {
            if (tocIndex.Value < 0 || tocIndex.Value >= tocFields.Count)
                throw new ArgumentException($"tocIndex must be between 0 and {tocFields.Count - 1}");
            tocFields[tocIndex.Value].Update();
        }
        else
        {
            foreach (var tocField in tocFields)
                tocField.Update();
        }

        doc.UpdateFields();
        doc.Save(outputPath);
        var updatedCount = tocIndex.HasValue ? 1 : tocFields.Count;
        return await Task.FromResult($"Updated {updatedCount} table of contents field(s): {outputPath}");
    }

    private async Task<string> AddIndex(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        SecurityHelper.ValidateFilePath(path, "path");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        var indexEntriesArray = arguments?["indexEntries"]?.AsArray() ?? throw new ArgumentException("indexEntries is required");
        var insertIndexAtEnd = arguments?["insertIndexAtEnd"]?.GetValue<bool?>() ?? true;
        var headingStyle = arguments?["headingStyle"]?.GetValue<string>() ?? "Heading 1";

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);

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
                        xeField += $" \\t \"{subEntry}\"";
                    if (!string.IsNullOrEmpty(pageRangeBookmark))
                        xeField += $" \\r \"{pageRangeBookmark}\"";
                    builder.InsertField(xeField);
                }
            }
        }

        if (insertIndexAtEnd)
        {
            builder.MoveToDocumentEnd();
            builder.InsertBreak(BreakType.PageBreak);
            builder.ParagraphFormat.Style = doc.Styles[headingStyle];
            builder.Writeln("Index");
            builder.ParagraphFormat.Style = doc.Styles["Normal"];
            builder.InsertField("INDEX \\e \" \" \\h \"A\"");
        }

        doc.Save(outputPath);
        return await Task.FromResult($"Index entries added. Total entries: {indexEntriesArray.Count}. Output: {outputPath}");
    }

    private async Task<string> AddCrossReference(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        SecurityHelper.ValidateFilePath(path, "path");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");
        var referenceType = arguments?["referenceType"]?.GetValue<string>() ?? throw new ArgumentException("referenceType is required");
        var referenceText = arguments?["referenceText"]?.GetValue<string>();
        var targetName = arguments?["targetName"]?.GetValue<string>() ?? throw new ArgumentException("targetName is required");
        var insertAsHyperlink = arguments?["insertAsHyperlink"]?.GetValue<bool?>() ?? true;
        var includeAboveBelow = arguments?["includeAboveBelow"]?.GetValue<bool?>() ?? false;

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        if (!string.IsNullOrEmpty(referenceText))
            builder.Write(referenceText);

        builder.InsertField($"REF {targetName} \\h");
        if (includeAboveBelow)
            builder.Write(" (above)");

        doc.Save(outputPath);
        return await Task.FromResult($"Cross-reference added: {outputPath}");
    }
}

