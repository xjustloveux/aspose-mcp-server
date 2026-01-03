using System.ComponentModel;
using System.Text;
using Aspose.Slides;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint text (add, edit, replace)
///     Merges: PptAddTextTool, PptEditTextTool, PptReplaceTextTool
/// </summary>
[McpServerToolType]
public class PptTextTool
{
    /// <summary>
    ///     Session manager for document session handling.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptTextTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory editing.</param>
    public PptTextTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "ppt_text")]
    [Description(@"Manage PowerPoint text. Supports 3 operations: add, edit, replace.
Searches text in AutoShapes, GroupShapes (recursive), and Table cells.

Coordinate unit: 1 inch = 72 points.

Usage examples:
- Add text: ppt_text(operation='add', path='presentation.pptx', slideIndex=0, text='Hello World', x=100, y=100)
- Edit text: ppt_text(operation='edit', path='presentation.pptx', slideIndex=0, shapeIndex=0, text='Updated Text')
- Replace text: ppt_text(operation='replace', path='presentation.pptx', findText='old', replaceText='new')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'add': Add text to slide (required params: path, slideIndex, text)
- 'edit': Edit text in shape (required params: path, slideIndex, shapeIndex, text)
- 'replace': Replace text in presentation (required params: path, findText, replaceText)")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (optional, defaults to input path)")]
        string? outputPath = null,
        [Description("Slide index (0-based, required for add/edit)")]
        int? slideIndex = null,
        [Description("Shape index (0-based, required for edit)")]
        int? shapeIndex = null,
        [Description("Text content (required for add/edit)")]
        string? text = null,
        [Description("Text to find (required for replace)")]
        string? findText = null,
        [Description("Text to replace with (required for replace)")]
        string? replaceText = null,
        [Description("Match case (optional, for replace, default: false)")]
        bool matchCase = false,
        [Description("X position in points (optional, for add, default: 50)")]
        float x = 50,
        [Description("Y position in points (optional, for add, default: 50)")]
        float y = 50,
        [Description("Text box width in points (optional, for add, default: 400)")]
        float width = 400,
        [Description("Text box height in points (optional, for add, default: 100)")]
        float height = 100)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path);

        return operation.ToLower() switch
        {
            "add" => AddText(ctx, outputPath, slideIndex, text, x, y, width, height),
            "edit" => EditText(ctx, outputPath, slideIndex, shapeIndex, text),
            "replace" => ReplaceText(ctx, outputPath, findText, replaceText, matchCase),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a text box to a slide.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The zero-based index of the slide.</param>
    /// <param name="text">The text content to add.</param>
    /// <param name="x">The X position in points.</param>
    /// <param name="y">The Y position in points.</param>
    /// <param name="width">The text box width in points.</param>
    /// <param name="height">The text box height in points.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when slideIndex or text is not provided.</exception>
    private static string AddText(DocumentContext<Presentation> ctx, string? outputPath, int? slideIndex,
        string? text, float x, float y, float width, float height)
    {
        if (!slideIndex.HasValue)
            throw new ArgumentException("slideIndex is required for add operation");
        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("text is required for add operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex.Value);

        var textBox = slide.Shapes.AddAutoShape(ShapeType.Rectangle, x, y, width, height);
        textBox.TextFrame.Text = text;
        textBox.FillFormat.FillType = FillType.NoFill;
        textBox.LineFormat.FillFormat.FillType = FillType.NoFill;

        ctx.Save(outputPath);

        return $"Text added to slide {slideIndex.Value}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Edits text in a shape on a slide.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The zero-based index of the slide.</param>
    /// <param name="shapeIndex">The zero-based index of the shape.</param>
    /// <param name="text">The new text content.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when slideIndex, shapeIndex, or text is not provided, or shape is not an
    ///     AutoShape.
    /// </exception>
    private static string EditText(DocumentContext<Presentation> ctx, string? outputPath, int? slideIndex,
        int? shapeIndex, string? text)
    {
        if (!slideIndex.HasValue)
            throw new ArgumentException("slideIndex is required for edit operation");
        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for edit operation");
        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("text is required for edit operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex.Value);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex.Value);

        if (shape is not IAutoShape autoShape)
            throw new ArgumentException(
                $"Shape at index {shapeIndex.Value} (Type: {shape.GetType().Name}) is not an AutoShape and cannot contain text");

        if (autoShape.TextFrame == null)
            autoShape.AddTextFrame("");

        if (autoShape.TextFrame != null)
        {
            autoShape.TextFrame.Paragraphs.Clear();
            var paragraph = new Paragraph();
            paragraph.Portions.Add(new Portion(text));
            autoShape.TextFrame.Paragraphs.Add(paragraph);
        }

        ctx.Save(outputPath);
        return
            $"Text updated on slide {slideIndex.Value}, shape {shapeIndex.Value}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Replaces text in the presentation across all shapes including GroupShapes and Tables.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="findText">The text to find.</param>
    /// <param name="replaceText">The text to replace with.</param>
    /// <param name="matchCase">True to match case, false for case-insensitive matching.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when findText or replaceText is not provided.</exception>
    private static string ReplaceText(DocumentContext<Presentation> ctx, string? outputPath, string? findText,
        string? replaceText, bool matchCase)
    {
        if (string.IsNullOrEmpty(findText))
            throw new ArgumentException("findText is required for replace operation");
        if (replaceText == null)
            throw new ArgumentException("replaceText is required for replace operation");

        var presentation = ctx.Document;
        var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var replacements = 0;

        foreach (var slide in presentation.Slides)
            replacements += ProcessShapesForReplace(slide.Shapes, findText, replaceText, comparison);

        ctx.Save(outputPath);
        return
            $"Replaced '{findText}' with '{replaceText}' ({replacements} occurrences). {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Recursively processes shapes for text replacement, including GroupShapes and Tables.
    /// </summary>
    /// <param name="shapes">The shape collection to process.</param>
    /// <param name="findText">The text to find.</param>
    /// <param name="replaceText">The text to replace with.</param>
    /// <param name="comparison">The string comparison type.</param>
    /// <returns>The number of replacements made.</returns>
    private static int ProcessShapesForReplace(IShapeCollection shapes, string findText, string replaceText,
        StringComparison comparison)
    {
        var replacements = 0;

        foreach (var shape in shapes)
            switch (shape)
            {
                case IAutoShape { TextFrame: not null } autoShape:
                    replacements += ReplaceInTextFrame(autoShape.TextFrame, findText, replaceText, comparison);
                    break;
                case IGroupShape groupShape:
                    replacements += ProcessShapesForReplace(groupShape.Shapes, findText, replaceText, comparison);
                    break;
                case ITable table:
                    replacements += ReplaceInTable(table, findText, replaceText, comparison);
                    break;
            }

        return replacements;
    }

    /// <summary>
    ///     Replaces text in a TextFrame while preserving formatting at the Portion level.
    /// </summary>
    /// <param name="textFrame">The text frame to process.</param>
    /// <param name="findText">The text to find.</param>
    /// <param name="replaceText">The text to replace with.</param>
    /// <param name="comparison">The string comparison type.</param>
    /// <returns>The number of replacements made (0 or 1).</returns>
    private static int ReplaceInTextFrame(ITextFrame textFrame, string findText, string replaceText,
        StringComparison comparison)
    {
        var originalText = textFrame.Text;
        if (string.IsNullOrEmpty(originalText)) return 0;

        if (originalText.IndexOf(findText, comparison) < 0) return 0;

        foreach (var para in textFrame.Paragraphs)
        foreach (var portion in para.Portions)
        {
            var portionText = portion.Text;
            if (string.IsNullOrEmpty(portionText)) continue;

            var newText = ReplaceAll(portionText, findText, replaceText, comparison);
            if (newText != portionText)
                portion.Text = newText;
        }

        return 1;
    }

    /// <summary>
    ///     Replaces text in all cells of a table.
    /// </summary>
    /// <param name="table">The table to process.</param>
    /// <param name="findText">The text to find.</param>
    /// <param name="replaceText">The text to replace with.</param>
    /// <param name="comparison">The string comparison type.</param>
    /// <returns>The number of replacements made.</returns>
    private static int ReplaceInTable(ITable table, string findText, string replaceText, StringComparison comparison)
    {
        var replacements = 0;

        for (var row = 0; row < table.Rows.Count; row++)
        for (var col = 0; col < table.Columns.Count; col++)
        {
            var cell = table[col, row];
            if (cell.TextFrame != null)
                replacements += ReplaceInTextFrame(cell.TextFrame, findText, replaceText, comparison);
        }

        return replacements;
    }

    /// <summary>
    ///     Replaces all occurrences of a string in source with replacement string.
    /// </summary>
    /// <param name="source">The source string to search in.</param>
    /// <param name="find">The text to find.</param>
    /// <param name="replace">The text to replace with.</param>
    /// <param name="comparison">The string comparison type.</param>
    /// <returns>The string with all occurrences replaced.</returns>
    private static string ReplaceAll(string source, string find, string replace, StringComparison comparison)
    {
        if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(find)) return source;

        var sb = new StringBuilder();
        var idx = 0;
        while (true)
        {
            var next = source.IndexOf(find, idx, comparison);
            if (next < 0)
            {
                sb.Append(source, idx, source.Length - idx);
                break;
            }

            sb.Append(source, idx, next - idx);
            sb.Append(replace);
            idx = next + find.Length;
        }

        return sb.ToString();
    }
}