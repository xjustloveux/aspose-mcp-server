using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordAddTextTool : IAsposeTool
{
    public string Description => "Add text to an existing Word document";

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
            text = new
            {
                type = "string",
                description = "Text to add"
            },
            fontName = new
            {
                type = "string",
                description = "Font name (optional, e.g., 'Arial', '微軟雅黑')"
            },
            fontSize = new
            {
                type = "number",
                description = "Font size (optional)"
            },
            bold = new
            {
                type = "boolean",
                description = "Bold text (optional)"
            },
            italic = new
            {
                type = "boolean",
                description = "Italic text (optional)"
            },
            color = new
            {
                type = "string",
                description = "Text color (hex format like 'FF0000' for red, or name like 'Red', 'Blue', optional)"
            }
        },
        required = new[] { "path", "text" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var text = arguments?["text"]?.GetValue<string>() ?? throw new ArgumentException("text is required");
        var fontName = arguments?["fontName"]?.GetValue<string>();
        var fontSize = arguments?["fontSize"]?.GetValue<double>();
        var bold = arguments?["bold"]?.GetValue<bool>() ?? false;
        var italic = arguments?["italic"]?.GetValue<bool>() ?? false;
        var color = arguments?["color"]?.GetValue<string>();

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        if (!string.IsNullOrEmpty(fontName))
            builder.Font.Name = fontName;
        if (fontSize.HasValue)
            builder.Font.Size = fontSize.Value;
        builder.Font.Bold = bold;
        builder.Font.Italic = italic;
        
        if (!string.IsNullOrEmpty(color))
        {
            try
            {
                if (color.StartsWith("#"))
                {
                    var hexColor = color.TrimStart('#');
                    var r = Convert.ToInt32(hexColor.Substring(0, 2), 16);
                    var g = Convert.ToInt32(hexColor.Substring(2, 2), 16);
                    var b = Convert.ToInt32(hexColor.Substring(4, 2), 16);
                    builder.Font.Color = System.Drawing.Color.FromArgb(r, g, b);
                }
                else
                {
                    builder.Font.Color = System.Drawing.Color.FromName(color);
                }
            }
            catch
            {
                // Ignore invalid color
            }
        }

        builder.Writeln(text);
        doc.Save(path);

        return await Task.FromResult($"Text added to document: {path}");
    }
}

