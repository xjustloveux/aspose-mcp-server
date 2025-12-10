using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordCreateStyleTool : IAsposeTool
{
    public string Description => "Create a custom reusable style in a Word document";

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
            styleName = new
            {
                type = "string",
                description = "Custom style name (e.g., 'MyHeading1')"
            },
            styleType = new
            {
                type = "string",
                description = "Style type: paragraph, character, table, list (default: paragraph)",
                @enum = new[] { "paragraph", "character", "table", "list" }
            },
            baseStyle = new
            {
                type = "string",
                description = "Base style to inherit from (e.g., 'Heading 1', 'Normal')"
            },
            fontName = new
            {
                type = "string",
                description = "Font name. If fontNameAscii and fontNameFarEast are provided, this will be used as fallback."
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
                description = "Font size in points"
            },
            bold = new
            {
                type = "boolean",
                description = "Bold text"
            },
            italic = new
            {
                type = "boolean",
                description = "Italic text"
            },
            underline = new
            {
                type = "boolean",
                description = "Underline text"
            },
            color = new
            {
                type = "string",
                description = "Text color (hex format like 'FF0000' or color name)"
            },
            alignment = new
            {
                type = "string",
                description = "Paragraph alignment: left, center, right, justify",
                @enum = new[] { "left", "center", "right", "justify" }
            },
            spaceBefore = new
            {
                type = "number",
                description = "Space before paragraph in points"
            },
            spaceAfter = new
            {
                type = "number",
                description = "Space after paragraph in points"
            },
            lineSpacing = new
            {
                type = "number",
                description = "Line spacing (1.0 = single, 1.5 = 1.5 lines, 2.0 = double)"
            }
        },
        required = new[] { "path", "styleName" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var styleName = arguments?["styleName"]?.GetValue<string>() ?? throw new ArgumentException("styleName is required");
        var styleTypeStr = arguments?["styleType"]?.GetValue<string>() ?? "paragraph";
        var baseStyle = arguments?["baseStyle"]?.GetValue<string>();
        var fontName = arguments?["fontName"]?.GetValue<string>();
        var fontNameAscii = arguments?["fontNameAscii"]?.GetValue<string>();
        var fontNameFarEast = arguments?["fontNameFarEast"]?.GetValue<string>();
        var fontSize = arguments?["fontSize"]?.GetValue<double?>();
        var bold = arguments?["bold"]?.GetValue<bool?>();
        var italic = arguments?["italic"]?.GetValue<bool?>();
        var underline = arguments?["underline"]?.GetValue<bool?>();
        var color = arguments?["color"]?.GetValue<string>();
        var alignment = arguments?["alignment"]?.GetValue<string>();
        var spaceBefore = arguments?["spaceBefore"]?.GetValue<double?>();
        var spaceAfter = arguments?["spaceAfter"]?.GetValue<double?>();
        var lineSpacing = arguments?["lineSpacing"]?.GetValue<double?>();

        var doc = new Document(path);

        // Check if style already exists
        if (doc.Styles[styleName] != null)
        {
            throw new InvalidOperationException($"樣式 '{styleName}' 已存在，請使用不同的名稱或先刪除現有樣式");
        }

        // Determine style type
        StyleType styleType = styleTypeStr.ToLower() switch
        {
            "character" => StyleType.Character,
            "table" => StyleType.Table,
            "list" => StyleType.List,
            _ => StyleType.Paragraph
        };

        // Create new style
        var style = doc.Styles.Add(styleType, styleName);

        // Set base style if provided
        if (!string.IsNullOrEmpty(baseStyle))
        {
            try
            {
                var baseStyleObj = doc.Styles[baseStyle];
                if (baseStyleObj != null)
                {
                    style.BaseStyleName = baseStyle;
                }
                else
                {
                    // Continue without base style - this is a warning, not an error
                }
            }
            catch
            {
                // Continue without base style
            }
        }

        // Apply font properties
        // Set font names (priority: fontNameAscii/fontNameFarEast > fontName)
        if (!string.IsNullOrEmpty(fontNameAscii))
            style.Font.NameAscii = fontNameAscii;
        
        if (!string.IsNullOrEmpty(fontNameFarEast))
            style.Font.NameFarEast = fontNameFarEast;
        
        if (!string.IsNullOrEmpty(fontName))
        {
            // If fontNameAscii/FarEast are not set, use fontName for both
            if (string.IsNullOrEmpty(fontNameAscii) && string.IsNullOrEmpty(fontNameFarEast))
            {
                style.Font.Name = fontName;
            }
            else
            {
                // If only one is set, use fontName as fallback for the other
                if (string.IsNullOrEmpty(fontNameAscii))
                    style.Font.NameAscii = fontName;
                if (string.IsNullOrEmpty(fontNameFarEast))
                    style.Font.NameFarEast = fontName;
            }
        }

        if (fontSize.HasValue)
            style.Font.Size = fontSize.Value;

        if (bold.HasValue)
            style.Font.Bold = bold.Value;

        if (italic.HasValue)
            style.Font.Italic = italic.Value;

        if (underline.HasValue)
            style.Font.Underline = underline.Value ? Underline.Single : Underline.None;

        if (!string.IsNullOrEmpty(color))
        {
            try
            {
                var parsedColor = ParseColor(color);
                style.Font.Color = parsedColor;
            }
            catch
            {
                // Ignore color parsing errors
            }
        }

        // Apply paragraph properties (only for paragraph and list styles)
        if (styleType == StyleType.Paragraph || styleType == StyleType.List)
        {
            if (!string.IsNullOrEmpty(alignment))
            {
                style.ParagraphFormat.Alignment = alignment.ToLower() switch
                {
                    "center" => ParagraphAlignment.Center,
                    "right" => ParagraphAlignment.Right,
                    "justify" => ParagraphAlignment.Justify,
                    _ => ParagraphAlignment.Left
                };
            }

            if (spaceBefore.HasValue)
                style.ParagraphFormat.SpaceBefore = spaceBefore.Value;

            if (spaceAfter.HasValue)
                style.ParagraphFormat.SpaceAfter = spaceAfter.Value;

            if (lineSpacing.HasValue)
            {
                style.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
                style.ParagraphFormat.LineSpacing = lineSpacing.Value * 12; // Convert to points
            }
        }

        doc.Save(outputPath);

        var result = $"成功創建自訂樣式: {styleName}\n";
        result += $"類型: {styleTypeStr}\n";
        if (!string.IsNullOrEmpty(baseStyle)) result += $"基於: {baseStyle}\n";
        if (!string.IsNullOrEmpty(fontNameAscii)) result += $"字體（英文）: {fontNameAscii}\n";
        if (!string.IsNullOrEmpty(fontNameFarEast)) result += $"字體（中文）: {fontNameFarEast}\n";
        if (!string.IsNullOrEmpty(fontName) && string.IsNullOrEmpty(fontNameAscii) && string.IsNullOrEmpty(fontNameFarEast)) 
            result += $"字體: {fontName}\n";
        if (fontSize.HasValue) result += $"字號: {fontSize.Value} pt\n";
        if (bold == true) result += "粗體: 是\n";
        if (italic == true) result += "斜體: 是\n";
        if (!string.IsNullOrEmpty(color)) result += $"顏色: {color}\n";
        result += $"輸出: {outputPath}\n";
        result += $"\n現在可以使用 word_add_text_with_style(styleName=\"{styleName}\") 來應用此樣式";

        return await Task.FromResult(result);
    }

    private System.Drawing.Color ParseColor(string color)
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
}

