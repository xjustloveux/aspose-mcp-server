using System.ComponentModel;
using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for paragraph operations in Word documents
///     Merges: WordInsertParagraphTool, WordDeleteParagraphTool, WordEditParagraphTool,
///     WordGetParagraphsTool, WordGetParagraphFormatTool, WordCopyParagraphFormatTool, WordMergeParagraphsTool
/// </summary>
[McpServerToolType]
public class WordParagraphTool
{
    /// <summary>
    ///     Handler registry for paragraph operations
    /// </summary>
    private readonly HandlerRegistry<Document> _handlerRegistry;

    /// <summary>
    ///     Identity accessor for session isolation
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document session operations
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the WordParagraphTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordParagraphTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Word.Paragraph");
    }

    /// <summary>
    ///     Executes a Word paragraph operation (insert, delete, edit, get, get_format, copy_format, merge).
    /// </summary>
    /// <param name="operation">The operation to perform: insert, delete, edit, get, get_format, copy_format, merge.</param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="paragraphIndex">Paragraph index (0-based, -1 for last paragraph).</param>
    /// <param name="text">Text content for the paragraph.</param>
    /// <param name="styleName">Style name to apply (e.g., 'Heading 1', 'Normal').</param>
    /// <param name="alignment">Text alignment: left, center, right, justify.</param>
    /// <param name="sectionIndex">Section index (0-based).</param>
    /// <param name="includeEmpty">Include empty paragraphs (default: true).</param>
    /// <param name="styleFilter">Filter by style name.</param>
    /// <param name="includeCommentParagraphs">Include paragraphs inside nested structures (default: true).</param>
    /// <param name="includeTextboxParagraphs">Include paragraphs inside TextBox/Shape objects (default: true).</param>
    /// <param name="includeRunDetails">Include detailed run-level formatting (default: true).</param>
    /// <param name="fontName">Font name.</param>
    /// <param name="fontNameAscii">Font name for ASCII characters.</param>
    /// <param name="fontNameFarEast">Font name for Far East characters.</param>
    /// <param name="fontSize">Font size in points.</param>
    /// <param name="bold">Bold text.</param>
    /// <param name="italic">Italic text.</param>
    /// <param name="underline">Underline text.</param>
    /// <param name="color">Text color hex.</param>
    /// <param name="indentLeft">Left indent in points.</param>
    /// <param name="indentRight">Right indent in points.</param>
    /// <param name="firstLineIndent">First line indent in points.</param>
    /// <param name="spaceBefore">Space before paragraph in points.</param>
    /// <param name="spaceAfter">Space after paragraph in points.</param>
    /// <param name="lineSpacing">Line spacing multiplier.</param>
    /// <param name="lineSpacingRule">Line spacing rule: single, oneAndHalf, double, atLeast, exactly, multiple.</param>
    /// <param name="tabStops">Custom tab stops array.</param>
    /// <param name="sourceParagraphIndex">Source paragraph index (for copy_format).</param>
    /// <param name="targetParagraphIndex">Target paragraph index (for copy_format).</param>
    /// <param name="startParagraphIndex">Start paragraph index (for merge).</param>
    /// <param name="endParagraphIndex">End paragraph index (for merge).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "word_paragraph")]
    [Description(
        @"Manage paragraphs in Word documents. Supports 7 operations: insert, delete, edit, get, get_format, copy_format, merge.

Usage examples:
- Insert paragraph: word_paragraph(operation='insert', path='doc.docx', paragraphIndex=0, text='New paragraph')
- Delete paragraph: word_paragraph(operation='delete', path='doc.docx', paragraphIndex=0)
- Edit format: word_paragraph(operation='edit', path='doc.docx', paragraphIndex=0, alignment='center', fontSize=14)
- Get paragraph: word_paragraph(operation='get', path='doc.docx', paragraphIndex=0)
- Get format: word_paragraph(operation='get_format', path='doc.docx', paragraphIndex=0)
- Copy format: word_paragraph(operation='copy_format', path='doc.docx', sourceParagraphIndex=0, targetParagraphIndex=1)
- Merge paragraphs: word_paragraph(operation='merge', path='doc.docx', startParagraphIndex=0, endParagraphIndex=2)

Important notes for 'get' operation:
- By default, returns ALL paragraphs in the document structure, including paragraphs inside Comment objects, table cells, and TextBoxes
- Use includeCommentParagraphs=false to get only Body paragraphs (visible in document body, excluding table cells and comments)
- Each paragraph shows its ParentNode type to help identify its location
- Paragraphs inside Comment objects are marked with '[Comment]' in the location field
- Paragraphs inside table cells are marked with '[Cell]' in the location field
- Paragraphs inside TextBoxes are marked with '[TextBox]' in the location field
- To check paragraph styles in table cells, use includeCommentParagraphs=true")]
    public string Execute(
        [Description("Operation: insert, delete, edit, get, get_format, copy_format, merge")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Paragraph index (0-based, -1 for last paragraph)")]
        int? paragraphIndex = null,
        [Description("Text content for the paragraph")]
        string? text = null,
        [Description("Style name to apply (e.g., 'Heading 1', 'Normal')")]
        string? styleName = null,
        [Description("Text alignment: left, center, right, justify")]
        string? alignment = null,
        [Description("Section index (0-based)")]
        int? sectionIndex = null,
        [Description("Include empty paragraphs (default: true)")]
        bool includeEmpty = true,
        [Description("Filter by style name")] string? styleFilter = null,
        [Description("Include paragraphs inside nested structures (default: true)")]
        bool includeCommentParagraphs = true,
        [Description("Include paragraphs inside TextBox/Shape objects (default: true)")]
        bool includeTextboxParagraphs = true,
        [Description("Include detailed run-level formatting (default: true)")]
        bool includeRunDetails = true,
        [Description("Font name")] string? fontName = null,
        [Description("Font name for ASCII characters")]
        string? fontNameAscii = null,
        [Description("Font name for Far East characters")]
        string? fontNameFarEast = null,
        [Description("Font size in points")] double? fontSize = null,
        [Description("Bold text")] bool? bold = null,
        [Description("Italic text")] bool? italic = null,
        [Description("Underline text")] bool? underline = null,
        [Description("Font color (hex format, e.g., '000000')")]
        string? color = null,
        [Description("Left indent in points")] double? indentLeft = null,
        [Description("Right indent in points")]
        double? indentRight = null,
        [Description("First line indent in points")]
        double? firstLineIndent = null,
        [Description("Space before paragraph in points")]
        double? spaceBefore = null,
        [Description("Space after paragraph in points")]
        double? spaceAfter = null,
        [Description("Line spacing value")] double? lineSpacing = null,
        [Description("Line spacing rule: single, oneAndHalf, double, atLeast, exactly, multiple")]
        string? lineSpacingRule = null,
        [Description("Custom tab stops array")]
        JsonArray? tabStops = null,
        [Description("Source paragraph index (for copy_format)")]
        int? sourceParagraphIndex = null,
        [Description("Target paragraph index (for copy_format)")]
        int? targetParagraphIndex = null,
        [Description("Start paragraph index (for merge)")]
        int? startParagraphIndex = null,
        [Description("End paragraph index (for merge)")]
        int? endParagraphIndex = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, paragraphIndex, text, styleName, alignment, sectionIndex,
            includeEmpty, styleFilter, includeCommentParagraphs, includeTextboxParagraphs, includeRunDetails,
            fontName, fontNameAscii, fontNameFarEast, fontSize, bold, italic, underline, color, indentLeft,
            indentRight, firstLineIndent, spaceBefore, spaceAfter, lineSpacing, lineSpacingRule, tabStops,
            sourceParagraphIndex, targetParagraphIndex, startParagraphIndex, endParagraphIndex);

        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Document>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);

        // Read-only operations don't need to save
        if (operation.ToLower() is "get" or "get_format")
            return result;

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return $"{result}\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    private static OperationParameters BuildParameters(
        string operation,
        int? paragraphIndex,
        string? text,
        string? styleName,
        string? alignment,
        int? sectionIndex,
        bool includeEmpty,
        string? styleFilter,
        bool includeCommentParagraphs,
        bool includeTextboxParagraphs,
        bool includeRunDetails,
        string? fontName,
        string? fontNameAscii,
        string? fontNameFarEast,
        double? fontSize,
        bool? bold,
        bool? italic,
        bool? underline,
        string? color,
        double? indentLeft,
        double? indentRight,
        double? firstLineIndent,
        double? spaceBefore,
        double? spaceAfter,
        double? lineSpacing,
        string? lineSpacingRule,
        JsonArray? tabStops,
        int? sourceParagraphIndex,
        int? targetParagraphIndex,
        int? startParagraphIndex,
        int? endParagraphIndex)
    {
        var parameters = new OperationParameters();

        switch (operation.ToLower())
        {
            case "insert":
                if (text != null) parameters.Set("text", text);
                if (paragraphIndex.HasValue) parameters.Set("paragraphIndex", paragraphIndex.Value);
                if (styleName != null) parameters.Set("styleName", styleName);
                if (alignment != null) parameters.Set("alignment", alignment);
                if (indentLeft.HasValue) parameters.Set("indentLeft", indentLeft.Value);
                if (indentRight.HasValue) parameters.Set("indentRight", indentRight.Value);
                if (firstLineIndent.HasValue) parameters.Set("firstLineIndent", firstLineIndent.Value);
                if (spaceBefore.HasValue) parameters.Set("spaceBefore", spaceBefore.Value);
                if (spaceAfter.HasValue) parameters.Set("spaceAfter", spaceAfter.Value);
                break;

            case "delete":
                if (paragraphIndex.HasValue) parameters.Set("paragraphIndex", paragraphIndex.Value);
                break;

            case "edit":
                if (paragraphIndex.HasValue) parameters.Set("paragraphIndex", paragraphIndex.Value);
                if (sectionIndex.HasValue) parameters.Set("sectionIndex", sectionIndex.Value);
                if (text != null) parameters.Set("text", text);
                if (styleName != null) parameters.Set("styleName", styleName);
                if (alignment != null) parameters.Set("alignment", alignment);
                if (fontName != null) parameters.Set("fontName", fontName);
                if (fontNameAscii != null) parameters.Set("fontNameAscii", fontNameAscii);
                if (fontNameFarEast != null) parameters.Set("fontNameFarEast", fontNameFarEast);
                if (fontSize.HasValue) parameters.Set("fontSize", fontSize.Value);
                if (bold.HasValue) parameters.Set("bold", bold.Value);
                if (italic.HasValue) parameters.Set("italic", italic.Value);
                if (underline.HasValue) parameters.Set("underline", underline.Value);
                if (color != null) parameters.Set("color", color);
                if (indentLeft.HasValue) parameters.Set("indentLeft", indentLeft.Value);
                if (indentRight.HasValue) parameters.Set("indentRight", indentRight.Value);
                if (firstLineIndent.HasValue) parameters.Set("firstLineIndent", firstLineIndent.Value);
                if (spaceBefore.HasValue) parameters.Set("spaceBefore", spaceBefore.Value);
                if (spaceAfter.HasValue) parameters.Set("spaceAfter", spaceAfter.Value);
                if (lineSpacing.HasValue) parameters.Set("lineSpacing", lineSpacing.Value);
                if (lineSpacingRule != null) parameters.Set("lineSpacingRule", lineSpacingRule);
                if (tabStops != null) parameters.Set("tabStops", tabStops);
                break;

            case "get":
                if (sectionIndex.HasValue) parameters.Set("sectionIndex", sectionIndex.Value);
                parameters.Set("includeEmpty", includeEmpty);
                if (styleFilter != null) parameters.Set("styleFilter", styleFilter);
                parameters.Set("includeCommentParagraphs", includeCommentParagraphs);
                parameters.Set("includeTextboxParagraphs", includeTextboxParagraphs);
                break;

            case "get_format":
                if (paragraphIndex.HasValue) parameters.Set("paragraphIndex", paragraphIndex.Value);
                parameters.Set("includeRunDetails", includeRunDetails);
                break;

            case "copy_format":
                if (sourceParagraphIndex.HasValue) parameters.Set("sourceParagraphIndex", sourceParagraphIndex.Value);
                if (targetParagraphIndex.HasValue) parameters.Set("targetParagraphIndex", targetParagraphIndex.Value);
                break;

            case "merge":
                if (startParagraphIndex.HasValue) parameters.Set("startParagraphIndex", startParagraphIndex.Value);
                if (endParagraphIndex.HasValue) parameters.Set("endParagraphIndex", endParagraphIndex.Value);
                break;
        }

        return parameters;
    }
}
