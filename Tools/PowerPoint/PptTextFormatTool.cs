using System.ComponentModel;
using System.Drawing;
using System.Text.Json;
using Aspose.Slides;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for PowerPoint text formatting (batch format text)
///     Merges: PptBatchFormatTextTool
/// </summary>
[McpServerToolType]
public class PptTextFormatTool
{
    /// <summary>
    ///     Session manager for document session handling.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptTextFormatTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory editing.</param>
    public PptTextFormatTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "ppt_text_format")]
    [Description(@"Batch format PowerPoint text. Formats font, size, bold, italic, color across slides.
Applies to text in AutoShapes and Table cells.

Color format: Hex color code (e.g., #FF5500, #RGB, #RRGGBB) or named colors (e.g., Red, Blue, DarkGreen).

Usage examples:
- Format all slides: ppt_text_format(path='presentation.pptx', fontName='Arial', fontSize=14, bold=true)
- Format specific slides: ppt_text_format(path='presentation.pptx', slideIndices='[0,1,2]', fontName='Times New Roman', fontSize=12)
- Format with color: ppt_text_format(path='presentation.pptx', color='#FF0000') or ppt_text_format(path='presentation.pptx', color='Red')")]
    public string Execute(
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (optional, defaults to input path)")]
        string? outputPath = null,
        [Description("Slide indices to apply as JSON array (optional; default all)")]
        string? slideIndices = null,
        [Description("Font name (optional)")] string? fontName = null,
        [Description("Font size (optional)")] double? fontSize = null,
        [Description("Bold (optional)")] bool? bold = null,
        [Description("Italic (optional)")] bool? italic = null,
        [Description("Text color: Hex (#FF5500, #RGB) or named color (Red, Blue, DarkGreen) (optional)")]
        string? color = null)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path);
        var presentation = ctx.Document;

        int[] targets;
        if (!string.IsNullOrWhiteSpace(slideIndices))
        {
            var indices = JsonSerializer.Deserialize<int[]>(slideIndices);
            targets = indices ?? Enumerable.Range(0, presentation.Slides.Count).ToArray();
        }
        else
        {
            targets = Enumerable.Range(0, presentation.Slides.Count).ToArray();
        }

        foreach (var idx in targets)
            if (idx < 0 || idx >= presentation.Slides.Count)
                throw new ArgumentException($"slide index {idx} out of range");

        Color? parsedColor = null;
        if (!string.IsNullOrWhiteSpace(color)) parsedColor = ColorHelper.ParseColor(color);

        var colorStr = parsedColor.HasValue
            ? $"#{parsedColor.Value.R:X2}{parsedColor.Value.G:X2}{parsedColor.Value.B:X2}"
            : null;

        foreach (var idx in targets)
        {
            var slide = presentation.Slides[idx];
            foreach (var shape in slide.Shapes)
                if (shape is IAutoShape { TextFrame: not null } auto)
                    ApplyFontToTextFrame(auto.TextFrame, fontName, fontSize, bold, italic, colorStr);
                else if (shape is ITable table)
                    ApplyFontToTable(table, fontName, fontSize, bold, italic, colorStr);
        }

        ctx.Save(outputPath);
        return $"Batch formatted text applied to {targets.Length} slides. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Applies font settings to all portions in a text frame.
    /// </summary>
    /// <param name="textFrame">The text frame to format.</param>
    /// <param name="fontName">Font name (null to skip).</param>
    /// <param name="fontSize">Font size (null to skip).</param>
    /// <param name="bold">Bold setting (null to skip).</param>
    /// <param name="italic">Italic setting (null to skip).</param>
    /// <param name="colorStr">Color string in hex format (null to skip).</param>
    private static void ApplyFontToTextFrame(ITextFrame textFrame, string? fontName, double? fontSize, bool? bold,
        bool? italic, string? colorStr)
    {
        foreach (var para in textFrame.Paragraphs)
        foreach (var portion in para.Portions)
            FontHelper.Ppt.ApplyFontSettings(portion.PortionFormat, fontName, fontSize, bold, italic, colorStr);
    }

    /// <summary>
    ///     Applies font settings to all cells in a table.
    /// </summary>
    /// <param name="table">The table to format.</param>
    /// <param name="fontName">Font name (null to skip).</param>
    /// <param name="fontSize">Font size (null to skip).</param>
    /// <param name="bold">Bold setting (null to skip).</param>
    /// <param name="italic">Italic setting (null to skip).</param>
    /// <param name="colorStr">Color string in hex format (null to skip).</param>
    private static void ApplyFontToTable(ITable table, string? fontName, double? fontSize, bool? bold, bool? italic,
        string? colorStr)
    {
        for (var row = 0; row < table.Rows.Count; row++)
        for (var col = 0; col < table.Columns.Count; col++)
        {
            var cell = table[col, row];
            if (cell.TextFrame != null)
                ApplyFontToTextFrame(cell.TextFrame, fontName, fontSize, bold, italic, colorStr);
        }
    }
}