using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Text;

namespace AsposeMcpServer.Tools;

public class PdfAddTextTool : IAsposeTool
{
    public string Description => "Add text to a PDF document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "PDF file path"
            },
            pageIndex = new
            {
                type = "number",
                description = "Page index (1-based)"
            },
            text = new
            {
                type = "string",
                description = "Text to add"
            },
            x = new
            {
                type = "number",
                description = "X position (optional, default: 100)"
            },
            y = new
            {
                type = "number",
                description = "Y position (optional, default: 700)"
            },
            fontName = new
            {
                type = "string",
                description = "Font name (optional, default: 'Arial', e.g., 'Arial', '微軟雅黑')"
            },
            fontSize = new
            {
                type = "number",
                description = "Font size (optional, default: 12)"
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
                description = "Text color (hex format like '#000000' for black, optional)"
            }
        },
        required = new[] { "path", "pageIndex", "text" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");
        var text = arguments?["text"]?.GetValue<string>() ?? throw new ArgumentException("text is required");
        var x = arguments?["x"]?.GetValue<double>() ?? 100;
        var y = arguments?["y"]?.GetValue<double>() ?? 700;
        var fontName = arguments?["fontName"]?.GetValue<string>() ?? "Arial";
        var fontSize = arguments?["fontSize"]?.GetValue<double>() ?? 12;
        var bold = arguments?["bold"]?.GetValue<bool>() ?? false;
        var italic = arguments?["italic"]?.GetValue<bool>() ?? false;
        var color = arguments?["color"]?.GetValue<string>();

        using var document = new Document(path);
        var page = document.Pages[pageIndex];

        var textFragment = new TextFragment(text);
        textFragment.Position = new Position(x, y);
        
        // Set font
        textFragment.TextState.Font = FontRepository.FindFont(fontName);
        textFragment.TextState.FontSize = (float)fontSize;
        textFragment.TextState.FontStyle = FontStyles.Regular;
        
        if (bold && italic)
            textFragment.TextState.FontStyle = FontStyles.Bold | FontStyles.Italic;
        else if (bold)
            textFragment.TextState.FontStyle = FontStyles.Bold;
        else if (italic)
            textFragment.TextState.FontStyle = FontStyles.Italic;
        
        // Set color
        if (!string.IsNullOrEmpty(color))
        {
            try
            {
                var hexColor = color.TrimStart('#');
                var r = Convert.ToInt32(hexColor.Length >= 2 ? hexColor.Substring(0, 2) : "00", 16);
                var g = Convert.ToInt32(hexColor.Length >= 4 ? hexColor.Substring(2, 2) : "00", 16);
                var b = Convert.ToInt32(hexColor.Length >= 6 ? hexColor.Substring(4, 2) : "00", 16);
                textFragment.TextState.ForegroundColor = Aspose.Pdf.Color.FromRgb(r, g, b);
            }
            catch
            {
                // Ignore invalid color, use default
            }
        }

        page.Paragraphs.Add(textFragment);
        document.Save(path);

        return await Task.FromResult($"Text added to page {pageIndex}: {path}");
    }
}

