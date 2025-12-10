using System.Text.Json.Nodes;
using System.Drawing;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptEditTableCellTool : IAsposeTool
{
    public string Description => "Edit a specific table cell with content and format";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Presentation file path"
            },
            slideIndex = new
            {
                type = "number",
                description = "Slide index (0-based)"
            },
            shapeIndex = new
            {
                type = "number",
                description = "Shape index of the table (0-based)"
            },
            rowIndex = new
            {
                type = "number",
                description = "Row index (0-based)"
            },
            columnIndex = new
            {
                type = "number",
                description = "Column index (0-based)"
            },
            text = new
            {
                type = "string",
                description = "Cell text content (optional)"
            },
            fillColor = new
            {
                type = "string",
                description = "Fill color hex, e.g. #FFAA00 (optional)"
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
            textColor = new
            {
                type = "string",
                description = "Text color hex (optional)"
            }
        },
        required = new[] { "path", "slideIndex", "shapeIndex", "rowIndex", "columnIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required");
        var rowIndex = arguments?["rowIndex"]?.GetValue<int>() ?? throw new ArgumentException("rowIndex is required");
        var columnIndex = arguments?["columnIndex"]?.GetValue<int>() ?? throw new ArgumentException("columnIndex is required");
        var text = arguments?["text"]?.GetValue<string>();
        var fillColor = arguments?["fillColor"]?.GetValue<string>();
        var fontName = arguments?["fontName"]?.GetValue<string>();
        var fontSize = arguments?["fontSize"]?.GetValue<float?>();
        var bold = arguments?["bold"]?.GetValue<bool?>();
        var italic = arguments?["italic"]?.GetValue<bool?>();
        var textColor = arguments?["textColor"]?.GetValue<string>();

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];
        if (shapeIndex < 0 || shapeIndex >= slide.Shapes.Count)
        {
            throw new ArgumentException($"shapeIndex must be between 0 and {slide.Shapes.Count - 1}");
        }

        var shape = slide.Shapes[shapeIndex];
        if (shape is not ITable table)
        {
            throw new ArgumentException($"Shape at index {shapeIndex} is not a table");
        }

        if (rowIndex < 0 || rowIndex >= table.Rows.Count)
        {
            throw new ArgumentException($"rowIndex must be between 0 and {table.Rows.Count - 1}");
        }
        if (columnIndex < 0 || columnIndex >= table.Columns.Count)
        {
            throw new ArgumentException($"columnIndex must be between 0 and {table.Columns.Count - 1}");
        }

        var cell = table[columnIndex, rowIndex];

        if (!string.IsNullOrEmpty(text))
        {
            cell.TextFrame.Text = text;
        }

        if (!string.IsNullOrWhiteSpace(fillColor))
        {
            var color = ColorTranslator.FromHtml(fillColor);
            cell.CellFormat.FillFormat.FillType = FillType.Solid;
            cell.CellFormat.FillFormat.SolidFillColor.Color = color;
        }

        if (cell.TextFrame.Paragraphs.Count > 0 && cell.TextFrame.Paragraphs[0].Portions.Count > 0)
        {
            var portion = cell.TextFrame.Paragraphs[0].Portions[0];
            if (!string.IsNullOrEmpty(fontName))
            {
                portion.PortionFormat.LatinFont = new FontData(fontName);
            }
            if (fontSize.HasValue)
            {
                portion.PortionFormat.FontHeight = fontSize.Value;
            }
            if (bold.HasValue)
            {
                portion.PortionFormat.FontBold = bold.Value ? NullableBool.True : NullableBool.False;
            }
            if (italic.HasValue)
            {
                portion.PortionFormat.FontItalic = italic.Value ? NullableBool.True : NullableBool.False;
            }
            if (!string.IsNullOrWhiteSpace(textColor))
            {
                var color = ColorTranslator.FromHtml(textColor);
                portion.PortionFormat.FillFormat.FillType = FillType.Solid;
                portion.PortionFormat.FillFormat.SolidFillColor.Color = color;
            }
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"Table cell [{rowIndex}, {columnIndex}] updated on slide {slideIndex}");
    }
}

