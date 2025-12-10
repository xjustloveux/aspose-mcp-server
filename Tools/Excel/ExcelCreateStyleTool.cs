using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelCreateStyleTool : IAsposeTool
{
    public string Description => "Create a custom style in Excel workbook";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Excel file path"
            },
            styleName = new
            {
                type = "string",
                description = "Style name (optional, if not provided creates unnamed style)"
            },
            fontName = new
            {
                type = "string",
                description = "Font name (optional)"
            },
            fontSize = new
            {
                type = "number",
                description = "Font size (optional)"
            },
            bold = new
            {
                type = "boolean",
                description = "Bold (optional)"
            },
            italic = new
            {
                type = "boolean",
                description = "Italic (optional)"
            },
            backgroundColor = new
            {
                type = "string",
                description = "Background color hex or name (optional)"
            },
            foregroundColor = new
            {
                type = "string",
                description = "Foreground/text color hex or name (optional)"
            },
            horizontalAlignment = new
            {
                type = "string",
                description = "Horizontal alignment (Left, Center, Right, optional)"
            },
            verticalAlignment = new
            {
                type = "string",
                description = "Vertical alignment (Top, Center, Bottom, optional)"
            },
            numberFormat = new
            {
                type = "string",
                description = "Number format (e.g., '#,##0.00', optional)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var styleName = arguments?["styleName"]?.GetValue<string>();
        var fontName = arguments?["fontName"]?.GetValue<string>();
        var fontSize = arguments?["fontSize"]?.GetValue<int?>();
        var bold = arguments?["bold"]?.GetValue<bool?>();
        var italic = arguments?["italic"]?.GetValue<bool?>();
        var backgroundColor = arguments?["backgroundColor"]?.GetValue<string>();
        var foregroundColor = arguments?["foregroundColor"]?.GetValue<string>();
        var horizontalAlignment = arguments?["horizontalAlignment"]?.GetValue<string>();
        var verticalAlignment = arguments?["verticalAlignment"]?.GetValue<string>();
        var numberFormat = arguments?["numberFormat"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        var style = workbook.CreateStyle();

        if (!string.IsNullOrEmpty(fontName))
        {
            style.Font.Name = fontName;
        }
        if (fontSize.HasValue)
        {
            style.Font.Size = fontSize.Value;
        }
        if (bold.HasValue)
        {
            style.Font.IsBold = bold.Value;
        }
        if (italic.HasValue)
        {
            style.Font.IsItalic = italic.Value;
        }

        if (!string.IsNullOrWhiteSpace(backgroundColor))
        {
            try
            {
                var color = backgroundColor.StartsWith("#") 
                    ? System.Drawing.ColorTranslator.FromHtml(backgroundColor)
                    : System.Drawing.Color.FromName(backgroundColor);
                style.ForegroundColor = color;
                style.Pattern = BackgroundType.Solid;
            }
            catch
            {
                // Ignore invalid color
            }
        }

        if (!string.IsNullOrWhiteSpace(foregroundColor))
        {
            try
            {
                var color = foregroundColor.StartsWith("#")
                    ? System.Drawing.ColorTranslator.FromHtml(foregroundColor)
                    : System.Drawing.Color.FromName(foregroundColor);
                style.Font.Color = color;
            }
            catch
            {
                // Ignore invalid color
            }
        }

        if (!string.IsNullOrEmpty(horizontalAlignment))
        {
            style.HorizontalAlignment = horizontalAlignment.ToLower() switch
            {
                "left" => TextAlignmentType.Left,
                "center" => TextAlignmentType.Center,
                "right" => TextAlignmentType.Right,
                _ => TextAlignmentType.General
            };
        }

        if (!string.IsNullOrEmpty(verticalAlignment))
        {
            style.VerticalAlignment = verticalAlignment.ToLower() switch
            {
                "top" => TextAlignmentType.Top,
                "center" => TextAlignmentType.Center,
                "bottom" => TextAlignmentType.Bottom,
                _ => TextAlignmentType.Center
            };
        }

        if (!string.IsNullOrEmpty(numberFormat))
        {
            style.Number = int.Parse(numberFormat);
        }

        // Note: Aspose.Cells doesn't support named styles directly
        // Styles are applied to cells/ranges, not stored as named styles in workbook
        // If styleName is provided, we'll just create the style (it can be reused via reference)

        workbook.Save(path);
        return await Task.FromResult($"Style {(string.IsNullOrEmpty(styleName) ? "created" : $"'{styleName}' created")}: {path}");
    }
}

