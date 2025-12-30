using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing watermarks in PDF documents
/// </summary>
public class PdfWatermarkTool : IAsposeTool
{
    public string Description => @"Manage watermarks in PDF documents. Supports 1 operation: add.

Usage examples:
- Add watermark: pdf_watermark(operation='add', path='doc.pdf', text='CONFIDENTIAL', fontSize=72, opacity=0.3)
- Add colored watermark: pdf_watermark(operation='add', path='doc.pdf', text='URGENT', color='Red')
- Add watermark to specific pages: pdf_watermark(operation='add', path='doc.pdf', text='DRAFT', pageRange='1,3,5-10')
- Add background watermark: pdf_watermark(operation='add', path='doc.pdf', text='SAMPLE', isBackground=true)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add text watermark (required params: path, text)",
                @enum = new[] { "add" }
            },
            path = new
            {
                type = "string",
                description = "PDF file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to overwrite input)"
            },
            text = new
            {
                type = "string",
                description = "Watermark text (required for add)"
            },
            opacity = new
            {
                type = "number",
                description = "Opacity (0.0 to 1.0, default: 0.3)",
                minimum = 0.0,
                maximum = 1.0
            },
            fontSize = new
            {
                type = "number",
                description = "Font size in points (default: 72)",
                minimum = 1
            },
            fontName = new
            {
                type = "string",
                description = "Font name (default: 'Arial')"
            },
            rotation = new
            {
                type = "number",
                description = "Rotation angle in degrees (default: 45)"
            },
            color = new
            {
                type = "string",
                description =
                    "Watermark color name (e.g., 'Red', 'Blue', 'Gray') or hex code (e.g., '#FF0000'). Default: 'Gray'"
            },
            pageRange = new
            {
                type = "string",
                description = "Page range to apply watermark (e.g., '1,3,5-10'). If not specified, applies to all pages"
            },
            isBackground = new
            {
                type = "boolean",
                description = "If true, watermark is placed behind text content. Default: false"
            },
            horizontalAlignment = new
            {
                type = "string",
                description = "Horizontal alignment (default: Center)",
                @enum = new[] { "Left", "Center", "Right" }
            },
            verticalAlignment = new
            {
                type = "string",
                description = "Vertical alignment (default: Center)",
                @enum = new[] { "Top", "Center", "Bottom" }
            }
        },
        required = new[] { "operation", "path", "text" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    /// <exception cref="ArgumentException">Thrown when operation is unknown or required parameters are missing</exception>
    /// <exception cref="FileNotFoundException">Thrown when input file does not exist</exception>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        return operation.ToLower() switch
        {
            "add" => await AddWatermark(path, outputPath, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a watermark to the PDF
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">
    ///     JSON arguments containing text, optional opacity, fontSize, fontName, rotation, color,
    ///     pageRange, isBackground, alignment
    /// </param>
    /// <returns>Success message indicating number of pages with watermark applied</returns>
    /// <exception cref="ArgumentException">Thrown when pageRange format is invalid</exception>
    private Task<string> AddWatermark(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var text = ArgumentHelper.GetString(arguments, "text");
            var opacity = ArgumentHelper.GetDouble(arguments, "opacity", "opacity", false, 0.3);
            var fontSize = ArgumentHelper.GetDouble(arguments, "fontSize", "fontSize", false, 72);
            var fontName = ArgumentHelper.GetString(arguments, "fontName", "Arial");
            var rotation = ArgumentHelper.GetDouble(arguments, "rotation", "rotation", false, 45);
            var colorName = ArgumentHelper.GetString(arguments, "color", "Gray");
            var pageRange = ArgumentHelper.GetStringNullable(arguments, "pageRange");
            var isBackground = ArgumentHelper.GetBool(arguments, "isBackground", false);
            var horizontalAlignment = ArgumentHelper.GetString(arguments, "horizontalAlignment", "Center");
            var verticalAlignment = ArgumentHelper.GetString(arguments, "verticalAlignment", "Center");

            using var document = new Document(path);

            var hAlign = horizontalAlignment.ToLower() switch
            {
                "left" => HorizontalAlignment.Left,
                "right" => HorizontalAlignment.Right,
                _ => HorizontalAlignment.Center
            };

            var vAlign = verticalAlignment.ToLower() switch
            {
                "top" => VerticalAlignment.Top,
                "bottom" => VerticalAlignment.Bottom,
                _ => VerticalAlignment.Center
            };

            var watermarkColor = ParseColor(colorName);
            var pageIndices = ParsePageRange(pageRange, document.Pages.Count);
            var appliedCount = 0;

            foreach (var pageIndex in pageIndices)
            {
                var page = document.Pages[pageIndex];
                var watermark = new WatermarkArtifact();
                var textState = new TextState
                {
                    ForegroundColor = watermarkColor
                };

                FontHelper.Pdf.ApplyFontSettings(textState, fontName, fontSize);

                watermark.SetTextAndState(text, textState);
                watermark.ArtifactHorizontalAlignment = hAlign;
                watermark.ArtifactVerticalAlignment = vAlign;
                watermark.Rotation = rotation;
                watermark.Opacity = opacity;
                watermark.IsBackground = isBackground;

                page.Artifacts.Add(watermark);
                appliedCount++;
            }

            document.Save(outputPath);

            return $"Watermark added to {appliedCount} page(s). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Parses a color string to Aspose Color
    /// </summary>
    /// <param name="colorName">Color name (e.g., 'Red', 'Blue') or hex code (e.g., '#FF0000')</param>
    /// <returns>Parsed Color object</returns>
    private static Color ParseColor(string colorName)
    {
        if (string.IsNullOrEmpty(colorName))
            return Color.Gray;

        if (colorName.StartsWith('#') && (colorName.Length == 7 || colorName.Length == 9))
            try
            {
                var hex = colorName.TrimStart('#');
                var r = Convert.ToByte(hex.Substring(0, 2), 16);
                var g = Convert.ToByte(hex.Substring(2, 2), 16);
                var b = Convert.ToByte(hex.Substring(4, 2), 16);
                return Color.FromRgb(r / 255.0, g / 255.0, b / 255.0);
            }
            catch
            {
                return Color.Gray;
            }

        return colorName.ToLower() switch
        {
            "red" => Color.Red,
            "blue" => Color.Blue,
            "green" => Color.Green,
            "black" => Color.Black,
            "white" => Color.White,
            "yellow" => Color.Yellow,
            "orange" => Color.Orange,
            "purple" => Color.Purple,
            "pink" => Color.Pink,
            "cyan" => Color.Cyan,
            "magenta" => Color.Magenta,
            "lightgray" => Color.LightGray,
            "darkgray" => Color.DarkGray,
            _ => Color.Gray
        };
    }

    /// <summary>
    ///     Parses a page range string into a list of page indices
    /// </summary>
    /// <param name="pageRange">Page range string (e.g., '1,3,5-10')</param>
    /// <param name="totalPages">Total number of pages in the document</param>
    /// <returns>List of 1-based page indices</returns>
    /// <exception cref="ArgumentException">Thrown when pageRange format is invalid</exception>
    private static List<int> ParsePageRange(string? pageRange, int totalPages)
    {
        if (string.IsNullOrEmpty(pageRange))
            return Enumerable.Range(1, totalPages).ToList();

        var result = new HashSet<int>();
        var parts = pageRange.Split(',', StringSplitOptions.RemoveEmptyEntries);

        foreach (var part in parts)
        {
            var trimmed = part.Trim();
            if (trimmed.Contains('-'))
            {
                var rangeParts = trimmed.Split('-');
                if (rangeParts.Length != 2 ||
                    !int.TryParse(rangeParts[0].Trim(), out var start) ||
                    !int.TryParse(rangeParts[1].Trim(), out var end))
                    throw new ArgumentException(
                        $"Invalid page range format: '{trimmed}'. Expected format: 'start-end' (e.g., '5-10')");

                if (start < 1 || end > totalPages || start > end)
                    throw new ArgumentException(
                        $"Page range '{trimmed}' is out of bounds. Document has {totalPages} page(s)");

                for (var i = start; i <= end; i++)
                    result.Add(i);
            }
            else
            {
                if (!int.TryParse(trimmed, out var pageNum))
                    throw new ArgumentException($"Invalid page number: '{trimmed}'");

                if (pageNum < 1 || pageNum > totalPages)
                    throw new ArgumentException(
                        $"Page number {pageNum} is out of bounds. Document has {totalPages} page(s)");

                result.Add(pageNum);
            }
        }

        return result.OrderBy(x => x).ToList();
    }
}