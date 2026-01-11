using System.ComponentModel;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Handlers.Word.Styles;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for managing styles in Word documents (get, create, apply, copy)
/// </summary>
[McpServerToolType]
public class WordStyleTool
{
    /// <summary>
    ///     Handler registry for style operations
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
    ///     Initializes a new instance of the WordStyleTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordStyleTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = WordStyleHandlerRegistry.Create();
    }

    /// <summary>
    ///     Executes a Word style operation (get_styles, create_style, apply_style, copy_styles).
    /// </summary>
    /// <param name="operation">The operation to perform: get_styles, create_style, apply_style, copy_styles.</param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="includeBuiltIn">Include built-in styles (for get_styles, default: false).</param>
    /// <param name="styleName">Style name (for create_style, apply_style).</param>
    /// <param name="styleType">Style type: paragraph, character, table, list (for create_style, default: paragraph).</param>
    /// <param name="baseStyle">Base style to inherit from (for create_style).</param>
    /// <param name="fontName">Font name (for create_style).</param>
    /// <param name="fontNameAscii">Font name for ASCII characters (for create_style).</param>
    /// <param name="fontNameFarEast">Font name for Far East characters (for create_style).</param>
    /// <param name="fontSize">Font size in points (for create_style).</param>
    /// <param name="bold">Bold text (for create_style).</param>
    /// <param name="italic">Italic text (for create_style).</param>
    /// <param name="underline">Underline text (for create_style).</param>
    /// <param name="color">Text color hex (for create_style).</param>
    /// <param name="alignment">Paragraph alignment: left, center, right, justify (for create_style).</param>
    /// <param name="spaceBefore">Space before paragraph in points (for create_style).</param>
    /// <param name="spaceAfter">Space after paragraph in points (for create_style).</param>
    /// <param name="lineSpacing">Line spacing multiplier (for create_style).</param>
    /// <param name="paragraphIndex">Paragraph index (0-based, for apply_style).</param>
    /// <param name="paragraphIndices">Array of paragraph indices (for apply_style).</param>
    /// <param name="sectionIndex">Section index (0-based, for apply_style, default: 0).</param>
    /// <param name="tableIndex">Table index (0-based, for apply_style).</param>
    /// <param name="applyToAllParagraphs">Apply to all paragraphs (for apply_style, default: false).</param>
    /// <param name="sourceDocument">Source document path to copy styles from (for copy_styles).</param>
    /// <param name="styleNames">Array of style names to copy (for copy_styles).</param>
    /// <param name="overwriteExisting">Overwrite existing styles (for copy_styles, default: false).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get_styles.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "word_style")]
    [Description(
        @"Manage styles in Word documents. Supports 4 operations: get_styles, create_style, apply_style, copy_styles.

Usage examples:
- Get styles: word_style(operation='get_styles', path='doc.docx', includeBuiltIn=true)
- Create style: word_style(operation='create_style', path='doc.docx', styleName='CustomStyle', styleType='paragraph', fontSize=14, bold=true)
- Apply style: word_style(operation='apply_style', path='doc.docx', styleName='Heading 1', paragraphIndex=0)
- Copy styles: word_style(operation='copy_styles', path='doc.docx', sourceDocument='template.docx')")]
    public string Execute(
        [Description("Operation: get_styles, create_style, apply_style, copy_styles")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Include built-in styles (for get_styles, default: false)")]
        bool includeBuiltIn = false,
        [Description("Style name (for create_style, apply_style)")]
        string? styleName = null,
        [Description("Style type: paragraph, character, table, list (for create_style, default: paragraph)")]
        string styleType = "paragraph",
        [Description("Base style to inherit from (for create_style)")]
        string? baseStyle = null,
        [Description("Font name (for create_style)")]
        string? fontName = null,
        [Description("Font name for ASCII characters (for create_style)")]
        string? fontNameAscii = null,
        [Description("Font name for Far East characters (for create_style)")]
        string? fontNameFarEast = null,
        [Description("Font size in points (for create_style)")]
        double? fontSize = null,
        [Description("Bold text (for create_style)")]
        bool? bold = null,
        [Description("Italic text (for create_style)")]
        bool? italic = null,
        [Description("Underline text (for create_style)")]
        bool? underline = null,
        [Description("Text color hex (for create_style)")]
        string? color = null,
        [Description("Paragraph alignment: left, center, right, justify (for create_style)")]
        string? alignment = null,
        [Description("Space before paragraph in points (for create_style)")]
        double? spaceBefore = null,
        [Description("Space after paragraph in points (for create_style)")]
        double? spaceAfter = null,
        [Description("Line spacing multiplier (for create_style)")]
        double? lineSpacing = null,
        [Description("Paragraph index (0-based, for apply_style)")]
        int? paragraphIndex = null,
        [Description("Array of paragraph indices (for apply_style)")]
        int[]? paragraphIndices = null,
        [Description("Section index (0-based, for apply_style, default: 0)")]
        int sectionIndex = 0,
        [Description("Table index (0-based, for apply_style)")]
        int? tableIndex = null,
        [Description("Apply to all paragraphs (for apply_style, default: false)")]
        bool applyToAllParagraphs = false,
        [Description("Source document path to copy styles from (for copy_styles)")]
        string? sourceDocument = null,
        [Description("Array of style names to copy (for copy_styles)")]
        string[]? styleNames = null,
        [Description("Overwrite existing styles (for copy_styles, default: false)")]
        bool overwriteExisting = false)
    {
        var parameters = BuildParameters(operation, includeBuiltIn, styleName, styleType, baseStyle, fontName,
            fontNameAscii, fontNameFarEast, fontSize, bold, italic, underline, color, alignment, spaceBefore,
            spaceAfter, lineSpacing, paragraphIndex, paragraphIndices, sectionIndex, tableIndex,
            applyToAllParagraphs, sourceDocument, styleNames, overwriteExisting);

        var handler = _handlerRegistry.GetHandler(operation);

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var effectiveOutputPath = outputPath ?? path;

        var operationContext = new OperationContext<Document>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = effectiveOutputPath
        };

        var result = handler.Execute(operationContext, parameters);

        if (operationContext.IsModified)
            ctx.Save(effectiveOutputPath);

        return ctx.IsSession ? result : $"{result}\n{ctx.GetOutputMessage(effectiveOutputPath)}";
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters.
    /// </summary>
    private static OperationParameters BuildParameters(
        string operation,
        bool includeBuiltIn,
        string? styleName,
        string styleType,
        string? baseStyle,
        string? fontName,
        string? fontNameAscii,
        string? fontNameFarEast,
        double? fontSize,
        bool? bold,
        bool? italic,
        bool? underline,
        string? color,
        string? alignment,
        double? spaceBefore,
        double? spaceAfter,
        double? lineSpacing,
        int? paragraphIndex,
        int[]? paragraphIndices,
        int sectionIndex,
        int? tableIndex,
        bool applyToAllParagraphs,
        string? sourceDocument,
        string[]? styleNames,
        bool overwriteExisting)
    {
        var parameters = new OperationParameters();

        switch (operation.ToLower())
        {
            case "get_styles":
                parameters.Set("includeBuiltIn", includeBuiltIn);
                break;

            case "create_style":
                if (styleName != null) parameters.Set("styleName", styleName);
                parameters.Set("styleType", styleType);
                if (baseStyle != null) parameters.Set("baseStyle", baseStyle);
                if (fontName != null) parameters.Set("fontName", fontName);
                if (fontNameAscii != null) parameters.Set("fontNameAscii", fontNameAscii);
                if (fontNameFarEast != null) parameters.Set("fontNameFarEast", fontNameFarEast);
                if (fontSize.HasValue) parameters.Set("fontSize", fontSize.Value);
                if (bold.HasValue) parameters.Set("bold", bold.Value);
                if (italic.HasValue) parameters.Set("italic", italic.Value);
                if (underline.HasValue) parameters.Set("underline", underline.Value);
                if (color != null) parameters.Set("color", color);
                if (alignment != null) parameters.Set("alignment", alignment);
                if (spaceBefore.HasValue) parameters.Set("spaceBefore", spaceBefore.Value);
                if (spaceAfter.HasValue) parameters.Set("spaceAfter", spaceAfter.Value);
                if (lineSpacing.HasValue) parameters.Set("lineSpacing", lineSpacing.Value);
                break;

            case "apply_style":
                if (styleName != null) parameters.Set("styleName", styleName);
                if (paragraphIndex.HasValue) parameters.Set("paragraphIndex", paragraphIndex.Value);
                if (paragraphIndices != null) parameters.Set("paragraphIndices", paragraphIndices);
                parameters.Set("sectionIndex", sectionIndex);
                if (tableIndex.HasValue) parameters.Set("tableIndex", tableIndex.Value);
                parameters.Set("applyToAllParagraphs", applyToAllParagraphs);
                break;

            case "copy_styles":
                if (sourceDocument != null) parameters.Set("sourceDocument", sourceDocument);
                if (styleNames != null) parameters.Set("styleNames", styleNames);
                parameters.Set("overwriteExisting", overwriteExisting);
                break;
        }

        return parameters;
    }
}
