using System.ComponentModel;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing watermarks in PDF documents
/// </summary>
[McpServerToolType]
public class PdfWatermarkTool
{
    /// <summary>
    ///     The session identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     The document session manager for managing in-memory document sessions.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PdfWatermarkTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public PdfWatermarkTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes a PDF watermark operation (add).
    /// </summary>
    /// <param name="operation">The operation to perform: add.</param>
    /// <param name="path">PDF file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="text">Watermark text (required for add).</param>
    /// <param name="opacity">Opacity (0.0 to 1.0).</param>
    /// <param name="fontSize">Font size in points.</param>
    /// <param name="fontName">Font name.</param>
    /// <param name="rotation">Rotation angle in degrees.</param>
    /// <param name="color">Watermark color name or hex code.</param>
    /// <param name="pageRange">Page range to apply watermark (e.g., '1,3,5-10').</param>
    /// <param name="isBackground">If true, watermark is placed behind text content.</param>
    /// <param name="horizontalAlignment">Horizontal alignment.</param>
    /// <param name="verticalAlignment">Vertical alignment.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "pdf_watermark")]
    [Description(@"Manage watermarks in PDF documents. Supports 1 operation: add.

Usage examples:
- Add watermark: pdf_watermark(operation='add', path='doc.pdf', text='CONFIDENTIAL', fontSize=72, opacity=0.3)
- Add colored watermark: pdf_watermark(operation='add', path='doc.pdf', text='URGENT', color='Red')
- Add watermark to specific pages: pdf_watermark(operation='add', path='doc.pdf', text='DRAFT', pageRange='1,3,5-10')
- Add background watermark: pdf_watermark(operation='add', path='doc.pdf', text='SAMPLE', isBackground=true)")]
    public string Execute(
        [Description("Operation: add")] string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Watermark text (required for add)")]
        string? text = null,
        [Description("Opacity (0.0 to 1.0, default: 0.3)")]
        double opacity = 0.3,
        [Description("Font size in points (default: 72)")]
        double fontSize = 72,
        [Description("Font name (default: 'Arial')")]
        string fontName = "Arial",
        [Description("Rotation angle in degrees (default: 45)")]
        double rotation = 45,
        [Description(
            "Watermark color name (e.g., 'Red', 'Blue', 'Gray') or hex code (e.g., '#FF0000'). Default: 'Gray'")]
        string color = "Gray",
        [Description("Page range to apply watermark (e.g., '1,3,5-10'). If not specified, applies to all pages")]
        string? pageRange = null,
        [Description("If true, watermark is placed behind text content. Default: false")]
        bool isBackground = false,
        [Description("Horizontal alignment (default: Center)")]
        string horizontalAlignment = "Center",
        [Description("Vertical alignment (default: Center)")]
        string verticalAlignment = "Center")
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLower() switch
        {
            "add" => AddWatermark(ctx, outputPath, text, opacity, fontSize, fontName, rotation, color, pageRange,
                isBackground, horizontalAlignment, verticalAlignment),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a watermark to the PDF document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="text">The watermark text.</param>
    /// <param name="opacity">The opacity of the watermark (0.0 to 1.0).</param>
    /// <param name="fontSize">The font size in points.</param>
    /// <param name="fontName">The font name.</param>
    /// <param name="rotation">The rotation angle in degrees.</param>
    /// <param name="colorName">The watermark color name or hex code.</param>
    /// <param name="pageRange">Optional page range to apply watermark to.</param>
    /// <param name="isBackground">Whether to place the watermark behind content.</param>
    /// <param name="horizontalAlignment">The horizontal alignment of the watermark.</param>
    /// <param name="verticalAlignment">The vertical alignment of the watermark.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when text is empty.</exception>
    private static string AddWatermark(DocumentContext<Document> ctx, string? outputPath, string? text, double opacity,
        double fontSize, string fontName, double rotation, string colorName, string? pageRange, bool isBackground,
        string horizontalAlignment, string verticalAlignment)
    {
        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("text is required for add operation");

        var document = ctx.Document;

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

        ctx.Save(outputPath);

        return $"Watermark added to {appliedCount} page(s). {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Parses a color name or hex code into a PDF Color object.
    /// </summary>
    /// <param name="colorName">The color name or hex code (e.g., "Red" or "#FF0000").</param>
    /// <returns>The parsed Color object, or Gray if parsing fails.</returns>
    private static Color ParseColor(string colorName)
    {
        if (string.IsNullOrEmpty(colorName))
            return Color.Gray;

        if (colorName.StartsWith('#') && (colorName.Length == 7 || colorName.Length == 9))
            try
            {
                var hex = colorName.TrimStart('#');
                var r = Convert.ToByte(hex[..2], 16);
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
    ///     Parses a page range string into a list of page indices.
    /// </summary>
    /// <param name="pageRange">The page range string (e.g., "1,3,5-10").</param>
    /// <param name="totalPages">The total number of pages in the document.</param>
    /// <returns>A list of 1-based page indices.</returns>
    /// <exception cref="ArgumentException">Thrown when the page range format is invalid or out of bounds.</exception>
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