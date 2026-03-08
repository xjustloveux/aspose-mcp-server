using System.ComponentModel;
using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for managing styles in Word documents (get, create, apply, copy)
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Word.Styles")]
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
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Word.Styles");
    }

    /// <summary>
    ///     Executes a Word style operation (list, create, apply, copy).
    /// </summary>
    /// <param name="operation">The operation to perform: list, create, apply, copy.</param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="includeBuiltIn">Include built-in styles (for list, default: false).</param>
    /// <param name="styleName">Style name (for create, apply).</param>
    /// <param name="styleType">Style type: paragraph, character, table, list (for create, default: paragraph).</param>
    /// <param name="baseStyle">Base style to inherit from (for create).</param>
    /// <param name="fontName">Font name (for create).</param>
    /// <param name="fontNameAscii">Font name for ASCII characters (for create).</param>
    /// <param name="fontNameFarEast">Font name for Far East characters (for create).</param>
    /// <param name="fontSize">Font size in points (for create).</param>
    /// <param name="bold">Bold text (for create).</param>
    /// <param name="italic">Italic text (for create).</param>
    /// <param name="underline">Underline text (for create).</param>
    /// <param name="color">Text color hex (for create).</param>
    /// <param name="alignment">Paragraph alignment: left, center, right, justify (for create).</param>
    /// <param name="spaceBefore">Space before paragraph in points (for create).</param>
    /// <param name="spaceAfter">Space after paragraph in points (for create).</param>
    /// <param name="lineSpacing">Line spacing multiplier (for create).</param>
    /// <param name="paragraphIndex">Paragraph index (0-based, for apply).</param>
    /// <param name="paragraphIndices">Array of paragraph indices (for apply).</param>
    /// <param name="sectionIndex">Section index (0-based, for apply, default: 0).</param>
    /// <param name="tableIndex">Table index (0-based, for apply).</param>
    /// <param name="applyToAllParagraphs">Apply to all paragraphs (for apply, default: false).</param>
    /// <param name="sourceDocument">Source document path to copy styles from (for copy).</param>
    /// <param name="styleNames">Array of style names to copy (for copy).</param>
    /// <param name="overwriteExisting">Overwrite existing styles (for copy, default: false).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for list.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "word_style",
        Title = "Word Style Operations",
        Destructive = false,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(
        @"Manage styles in Word documents. Supports 4 operations: list, create, apply, copy.

Usage examples:
- Get styles: word_style(operation='list', path='doc.docx', includeBuiltIn=true)
- Create style: word_style(operation='create', path='doc.docx', styleName='CustomStyle', styleType='paragraph', fontSize=14, bold=true)
- Apply style: word_style(operation='apply', path='doc.docx', styleName='Heading 1', paragraphIndex=0)
- Copy styles: word_style(operation='copy', path='doc.docx', sourceDocument='template.docx')")]
    public object Execute(
        [Description("Operation: list, create, apply, copy")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Include built-in styles (for list, default: false)")]
        bool includeBuiltIn = false,
        [Description("Style name (for create, apply)")]
        string? styleName = null,
        [Description("Style type: paragraph, character, table, list (for create, default: paragraph)")]
        string styleType = "paragraph",
        [Description("Base style to inherit from (for create)")]
        string? baseStyle = null,
        [Description("Font name (for create)")]
        string? fontName = null,
        [Description("Font name for ASCII characters (for create)")]
        string? fontNameAscii = null,
        [Description("Font name for Far East characters (for create)")]
        string? fontNameFarEast = null,
        [Description("Font size in points (for create)")]
        double? fontSize = null,
        [Description("Bold text (for create)")]
        bool? bold = null,
        [Description("Italic text (for create)")]
        bool? italic = null,
        [Description("Underline text (for create)")]
        bool? underline = null,
        [Description("Text color hex (for create)")]
        string? color = null,
        [Description("Paragraph alignment: left, center, right, justify (for create)")]
        string? alignment = null,
        [Description("Space before paragraph in points (for create)")]
        double? spaceBefore = null,
        [Description("Space after paragraph in points (for create)")]
        double? spaceAfter = null,
        [Description("Line spacing multiplier (for create)")]
        double? lineSpacing = null,
        [Description("Paragraph index (0-based, for apply)")]
        int? paragraphIndex = null,
        [Description("Array of paragraph indices (for apply)")]
        int[]? paragraphIndices = null,
        [Description("Section index (0-based, for apply, default: 0)")]
        int sectionIndex = 0,
        [Description("Table index (0-based, for apply)")]
        int? tableIndex = null,
        [Description("Apply to all paragraphs (for apply, default: false)")]
        bool applyToAllParagraphs = false,
        [Description("Source document path to copy styles from (for copy)")]
        string? sourceDocument = null,
        [Description("Array of style names to copy (for copy)")]
        string[]? styleNames = null,
        [Description("Overwrite existing styles (for copy, default: false)")]
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

        return ResultHelper.FinalizeResult((dynamic)result, ctx, effectiveOutputPath);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
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
        return operation.ToLower() switch
        {
            "list" => BuildGetStylesParameters(includeBuiltIn),
            "create" => BuildCreateStyleParameters(styleName, styleType, baseStyle, fontName, fontNameAscii,
                fontNameFarEast, fontSize, bold, italic, underline, color, alignment, spaceBefore, spaceAfter,
                lineSpacing),
            "apply" => BuildApplyStyleParameters(styleName, paragraphIndex, paragraphIndices, sectionIndex,
                tableIndex, applyToAllParagraphs),
            "copy" => BuildCopyStylesParameters(sourceDocument, styleNames, overwriteExisting),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the list operation.
    /// </summary>
    /// <param name="includeBuiltIn">Whether to include built-in styles.</param>
    /// <returns>OperationParameters configured for getting styles.</returns>
    private static OperationParameters BuildGetStylesParameters(bool includeBuiltIn)
    {
        var parameters = new OperationParameters();
        parameters.Set("includeBuiltIn", includeBuiltIn);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the create operation.
    /// </summary>
    /// <param name="styleName">The style name.</param>
    /// <param name="styleType">The style type: paragraph, character, table, list.</param>
    /// <param name="baseStyle">The base style to inherit from.</param>
    /// <param name="fontName">The font name.</param>
    /// <param name="fontNameAscii">The font name for ASCII characters.</param>
    /// <param name="fontNameFarEast">The font name for Far East characters.</param>
    /// <param name="fontSize">The font size in points.</param>
    /// <param name="bold">Whether the text is bold.</param>
    /// <param name="italic">Whether the text is italic.</param>
    /// <param name="underline">Whether the text is underlined.</param>
    /// <param name="color">The text color in hex format.</param>
    /// <param name="alignment">The paragraph alignment: left, center, right, justify.</param>
    /// <param name="spaceBefore">The space before paragraph in points.</param>
    /// <param name="spaceAfter">The space after paragraph in points.</param>
    /// <param name="lineSpacing">The line spacing multiplier.</param>
    /// <returns>OperationParameters configured for creating a style.</returns>
    private static OperationParameters BuildCreateStyleParameters(string? styleName, string styleType,
        string? baseStyle,
        string? fontName, string? fontNameAscii, string? fontNameFarEast, double? fontSize, bool? bold, bool? italic,
        bool? underline, string? color, string? alignment, double? spaceBefore, double? spaceAfter, double? lineSpacing)
    {
        var parameters = new OperationParameters();
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
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the apply operation.
    /// </summary>
    /// <param name="styleName">The style name to apply.</param>
    /// <param name="paragraphIndex">The paragraph index (0-based).</param>
    /// <param name="paragraphIndices">The array of paragraph indices.</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <param name="tableIndex">The table index (0-based).</param>
    /// <param name="applyToAllParagraphs">Whether to apply to all paragraphs.</param>
    /// <returns>OperationParameters configured for applying a style.</returns>
    private static OperationParameters BuildApplyStyleParameters(string? styleName, int? paragraphIndex,
        int[]? paragraphIndices, int sectionIndex, int? tableIndex, bool applyToAllParagraphs)
    {
        var parameters = new OperationParameters();
        if (styleName != null) parameters.Set("styleName", styleName);
        if (paragraphIndex.HasValue) parameters.Set("paragraphIndex", paragraphIndex.Value);
        if (paragraphIndices != null) parameters.Set("paragraphIndices", paragraphIndices);
        parameters.Set("sectionIndex", sectionIndex);
        if (tableIndex.HasValue) parameters.Set("tableIndex", tableIndex.Value);
        parameters.Set("applyToAllParagraphs", applyToAllParagraphs);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the copy operation.
    /// </summary>
    /// <param name="sourceDocument">The source document path to copy styles from.</param>
    /// <param name="styleNames">The array of style names to copy.</param>
    /// <param name="overwriteExisting">Whether to overwrite existing styles.</param>
    /// <returns>OperationParameters configured for copying styles.</returns>
    private static OperationParameters BuildCopyStylesParameters(string? sourceDocument, string[]? styleNames,
        bool overwriteExisting)
    {
        var parameters = new OperationParameters();
        if (sourceDocument != null) parameters.Set("sourceDocument", sourceDocument);
        if (styleNames != null) parameters.Set("styleNames", styleNames);
        parameters.Set("overwriteExisting", overwriteExisting);
        return parameters;
    }
}
