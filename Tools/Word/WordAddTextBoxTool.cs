using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeMcpServer.Tools;

public class WordAddTextBoxTool : IAsposeTool
{
    public string Description => "Add a text box with custom styling to a Word document";

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
            text = new
            {
                type = "string",
                description = "Text content"
            },
            width = new
            {
                type = "number",
                description = "Width in points (default: 200)"
            },
            height = new
            {
                type = "number",
                description = "Height in points (default: 100)"
            },
            positionX = new
            {
                type = "number",
                description = "Horizontal position in points from left (default: 100)"
            },
            positionY = new
            {
                type = "number",
                description = "Vertical position in points from top (default: 100)"
            },
            backgroundColor = new
            {
                type = "string",
                description = "Background color (hex format, e.g., 'fff2cc')"
            },
            borderColor = new
            {
                type = "string",
                description = "Border color (hex format, e.g., 'ffc000')"
            },
            borderWidth = new
            {
                type = "number",
                description = "Border width in points (default: 1)"
            },
            fontName = new
            {
                type = "string",
                description = "Font name (optional). If fontNameAscii and fontNameFarEast are provided, this will be used as fallback."
            },
            fontNameAscii = new
            {
                type = "string",
                description = "Font name for ASCII characters (English, e.g., 'Times New Roman')"
            },
            fontNameFarEast = new
            {
                type = "string",
                description = "Font name for Far East characters (Chinese/Japanese/Korean, e.g., '標楷體')"
            },
            fontSize = new
            {
                type = "number",
                description = "Font size in points (optional)"
            },
            bold = new
            {
                type = "boolean",
                description = "Bold text (optional)"
            },
            textAlignment = new
            {
                type = "string",
                description = "Text alignment: left, center, right (default: left)",
                @enum = new[] { "left", "center", "right" }
            }
        },
        required = new[] { "path", "text" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var text = arguments?["text"]?.GetValue<string>() ?? throw new ArgumentException("text is required");
        var width = arguments?["width"]?.GetValue<double>() ?? 200;
        var height = arguments?["height"]?.GetValue<double>() ?? 100;
        var positionX = arguments?["positionX"]?.GetValue<double>() ?? 100;
        var positionY = arguments?["positionY"]?.GetValue<double>() ?? 100;
        var backgroundColor = arguments?["backgroundColor"]?.GetValue<string>();
        var borderColor = arguments?["borderColor"]?.GetValue<string>();
        var borderWidth = arguments?["borderWidth"]?.GetValue<double>() ?? 1;
        var fontName = arguments?["fontName"]?.GetValue<string>();
        var fontNameAscii = arguments?["fontNameAscii"]?.GetValue<string>();
        var fontNameFarEast = arguments?["fontNameFarEast"]?.GetValue<string>();
        var fontSize = arguments?["fontSize"]?.GetValue<double?>();
        var bold = arguments?["bold"]?.GetValue<bool?>();
        var textAlignment = arguments?["textAlignment"]?.GetValue<string>() ?? "left";

        var doc = new Document(path);
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        // Create text box
        var textBox = new Shape(doc, ShapeType.TextBox);
        textBox.Width = width;
        textBox.Height = height;
        textBox.Left = positionX;
        textBox.Top = positionY;
        textBox.WrapType = WrapType.None;
        textBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        textBox.RelativeVerticalPosition = RelativeVerticalPosition.Page;

        // Set background color
        if (!string.IsNullOrEmpty(backgroundColor))
        {
            textBox.Fill.Color = ParseColor(backgroundColor);
            textBox.Fill.Visible = true;
        }

        // Set border
        if (!string.IsNullOrEmpty(borderColor))
        {
            textBox.Stroke.Color = ParseColor(borderColor);
            textBox.Stroke.Weight = borderWidth;
            textBox.Stroke.Visible = true;
        }

        // Add text to text box
        var para = new Paragraph(doc);
        var run = new Run(doc, text);

        // Apply text formatting
        // Set font names (priority: fontNameAscii/fontNameFarEast > fontName)
        if (!string.IsNullOrEmpty(fontNameAscii))
            run.Font.NameAscii = fontNameAscii;
        
        if (!string.IsNullOrEmpty(fontNameFarEast))
            run.Font.NameFarEast = fontNameFarEast;
        
        if (!string.IsNullOrEmpty(fontName))
        {
            // If fontNameAscii/FarEast are not set, use fontName for both
            if (string.IsNullOrEmpty(fontNameAscii) && string.IsNullOrEmpty(fontNameFarEast))
            {
                run.Font.Name = fontName;
            }
            else
            {
                // If only one is set, use fontName as fallback for the other
                if (string.IsNullOrEmpty(fontNameAscii))
                    run.Font.NameAscii = fontName;
                if (string.IsNullOrEmpty(fontNameFarEast))
                    run.Font.NameFarEast = fontName;
            }
        }

        if (fontSize.HasValue)
            run.Font.Size = fontSize.Value;

        if (bold.HasValue)
            run.Font.Bold = bold.Value;

        // Set paragraph alignment
        para.ParagraphFormat.Alignment = textAlignment.ToLower() switch
        {
            "center" => ParagraphAlignment.Center,
            "right" => ParagraphAlignment.Right,
            _ => ParagraphAlignment.Left
        };

        para.AppendChild(run);
        textBox.AppendChild(para);

        // Insert text box
        builder.InsertNode(textBox);

        doc.Save(outputPath);

        var result = $"成功添加文字框\n";
        result += $"尺寸: {width} x {height} pt\n";
        result += $"位置: ({positionX}, {positionY})\n";
        if (!string.IsNullOrEmpty(backgroundColor)) result += $"背景色: {backgroundColor}\n";
        if (!string.IsNullOrEmpty(borderColor)) result += $"邊框色: {borderColor}\n";
        result += $"輸出: {outputPath}";

        return await Task.FromResult(result);
    }

    private System.Drawing.Color ParseColor(string color)
    {
        try
        {
            if (color.StartsWith("#"))
                color = color.Substring(1);

            if (color.Length == 6)
            {
                int r = Convert.ToInt32(color.Substring(0, 2), 16);
                int g = Convert.ToInt32(color.Substring(2, 2), 16);
                int b = Convert.ToInt32(color.Substring(4, 2), 16);
                return System.Drawing.Color.FromArgb(r, g, b);
            }
            else
            {
                return System.Drawing.Color.FromName(color);
            }
        }
        catch
        {
            return System.Drawing.Color.White;
        }
    }
}

