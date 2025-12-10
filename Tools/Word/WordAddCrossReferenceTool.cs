using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Fields;

namespace AsposeMcpServer.Tools;

public class WordAddCrossReferenceTool : IAsposeTool
{
    public string Description => "Add cross-reference (to heading, bookmark, figure, table, etc.) in Word document";

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
            referenceType = new
            {
                type = "string",
                description = "Reference type: 'Heading', 'Bookmark', 'Figure', 'Table', 'Equation'",
                @enum = new[] { "Heading", "Bookmark", "Figure", "Table", "Equation" }
            },
            referenceText = new
            {
                type = "string",
                description = "Text to insert before reference (e.g., 'See ', optional)"
            },
            targetName = new
            {
                type = "string",
                description = "Target name (heading text, bookmark name, etc.)"
            },
            insertAsHyperlink = new
            {
                type = "boolean",
                description = "Insert as hyperlink (optional, default: true)"
            },
            includeAboveBelow = new
            {
                type = "boolean",
                description = "Include 'above' or 'below' text (optional, default: false)"
            }
        },
        required = new[] { "path", "referenceType", "targetName" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var referenceType = arguments?["referenceType"]?.GetValue<string>() ?? throw new ArgumentException("referenceType is required");
        var referenceText = arguments?["referenceText"]?.GetValue<string>();
        var targetName = arguments?["targetName"]?.GetValue<string>() ?? throw new ArgumentException("targetName is required");
        var insertAsHyperlink = arguments?["insertAsHyperlink"]?.GetValue<bool?>() ?? true;
        var includeAboveBelow = arguments?["includeAboveBelow"]?.GetValue<bool?>() ?? false;

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        if (!string.IsNullOrEmpty(referenceText))
        {
            builder.Write(referenceText);
        }

        // Insert cross-reference field
        var fieldType = referenceType.ToLower() switch
        {
            "heading" => FieldType.FieldRef,
            "bookmark" => FieldType.FieldRef,
            "figure" => FieldType.FieldRef,
            "table" => FieldType.FieldRef,
            "equation" => FieldType.FieldRef,
            _ => FieldType.FieldRef
        };

        builder.InsertField($"REF {targetName} \\h");
        if (includeAboveBelow)
        {
            builder.Write(" (above)");
        }

        doc.Save(path);
        return await Task.FromResult($"Cross-reference added: {path}");
    }
}

