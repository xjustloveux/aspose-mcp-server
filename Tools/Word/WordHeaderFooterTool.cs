using System.ComponentModel;
using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Handlers.Word.HeaderFooter;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for header and footer operations in Word documents
///     Merges: WordSetHeaderTextTool, WordSetFooterTextTool, WordSetHeaderImageTool, WordSetFooterImageTool,
///     WordSetHeaderLineTool, WordSetFooterLineTool, WordSetHeaderTabStopsTool, WordSetFooterTabStopsTool,
///     WordSetHeaderFooterTool, WordGetHeadersFootersTool
/// </summary>
[McpServerToolType]
public class WordHeaderFooterTool
{
    /// <summary>
    ///     Handler registry for header/footer operations
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
    ///     Initializes a new instance of the WordHeaderFooterTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordHeaderFooterTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = WordHeaderFooterHandlerRegistry.Create();
    }

    /// <summary>
    ///     Executes a Word header/footer operation (set_header_text, set_footer_text, set_header_image, set_footer_image,
    ///     set_header_line, set_footer_line, set_header_tabs, set_footer_tabs, set_header_footer, get).
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: set_header_text, set_footer_text, set_header_image, set_footer_image,
    ///     set_header_line, set_footer_line, set_header_tabs, set_footer_tabs, set_header_footer, get.
    /// </param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="headerLeft">Header left section text (for set_header_text).</param>
    /// <param name="headerCenter">Header center section text (for set_header_text).</param>
    /// <param name="headerRight">Header right section text (for set_header_text).</param>
    /// <param name="footerLeft">Footer left section text (for set_footer_text).</param>
    /// <param name="footerCenter">Footer center section text (for set_footer_text).</param>
    /// <param name="footerRight">Footer right section text (for set_footer_text).</param>
    /// <param name="imagePath">Path to image file (for set_header_image/set_footer_image).</param>
    /// <param name="alignment">Image alignment: left, center, right (for image operations).</param>
    /// <param name="imageWidth">Image width in points (for image operations).</param>
    /// <param name="imageHeight">Image height in points (for image operations).</param>
    /// <param name="lineStyle">Line style: single, double, thick (for line operations).</param>
    /// <param name="lineWidth">Line width in points (for line operations).</param>
    /// <param name="tabStops">Tab stops array (for tab operations).</param>
    /// <param name="fontName">Font name.</param>
    /// <param name="fontNameAscii">Font name for ASCII characters.</param>
    /// <param name="fontNameFarEast">Font name for Far East characters.</param>
    /// <param name="fontSize">Font size in points.</param>
    /// <param name="sectionIndex">Section index (0-based).</param>
    /// <param name="headerFooterType">Header/footer type: Primary, FirstPage, EvenPage.</param>
    /// <param name="isFloating">Make image floating instead of inline.</param>
    /// <param name="autoTabStops">Automatically add tab stops when using left/center/right text.</param>
    /// <param name="clearExisting">Clear existing content before setting new content.</param>
    /// <param name="clearTextOnly">Only clear text content, preserve images and shapes.</param>
    /// <param name="removeExisting">Remove existing images before adding new one.</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "word_header_footer")]
    [Description(
        @"Manage headers and footers in Word documents. Supports 10 operations: set_header_text, set_footer_text, set_header_image, set_footer_image, set_header_line, set_footer_line, set_header_tabs, set_footer_tabs, set_header_footer, get.

Usage examples:
- Set header text: word_header_footer(operation='set_header_text', path='doc.docx', headerLeft='Left', headerCenter='Center', headerRight='Right')
- Set footer text: word_header_footer(operation='set_footer_text', path='doc.docx', footerLeft='Page', footerCenter='', footerRight='{PAGE}')
- Set header image: word_header_footer(operation='set_header_image', path='doc.docx', imagePath='logo.png')
- Get headers/footers: word_header_footer(operation='get', path='doc.docx')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'set_header_text': Set header text (required params: path)
- 'set_footer_text': Set footer text (required params: path)
- 'set_header_image': Set header image (required params: path, imagePath)
- 'set_footer_image': Set footer image (required params: path, imagePath)
- 'set_header_line': Set header line (required params: path)
- 'set_footer_line': Set footer line (required params: path)
- 'set_header_tabs': Set header tab stops (required params: path)
- 'set_footer_tabs': Set footer tab stops (required params: path)
- 'set_header_footer': Set header and footer together (required params: path)
- 'get': Get headers and footers info (required params: path)")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Header left section text (optional, for set_header_text operation)")]
        string? headerLeft = null,
        [Description("Header center section text (optional, for set_header_text operation)")]
        string? headerCenter = null,
        [Description("Header right section text (optional, for set_header_text operation)")]
        string? headerRight = null,
        [Description("Footer left section text (optional, for set_footer_text operation)")]
        string? footerLeft = null,
        [Description("Footer center section text (optional, for set_footer_text operation)")]
        string? footerCenter = null,
        [Description("Footer right section text (optional, for set_footer_text operation)")]
        string? footerRight = null,
        [Description("Path to image file (required for set_header_image/set_footer_image operations)")]
        string? imagePath = null,
        [Description("Image alignment: left, center, right (optional, default: left, for image operations)")]
        string alignment = "left",
        [Description("Image width in points (optional, default: 50, for image operations)")]
        double? imageWidth = null,
        [Description(
            "Image height in points (optional, maintains aspect ratio if not specified, for image operations)")]
        double? imageHeight = null,
        [Description("Line style: single, double, thick (optional, for line operations)")]
        string lineStyle = "single",
        [Description("Line width in points (optional, for line operations)")]
        double? lineWidth = null,
        [Description("Tab stops (optional, for tab operations)")]
        JsonArray? tabStops = null,
        [Description("Font name (optional)")] string? fontName = null,
        [Description("Font name for ASCII characters (English, optional)")]
        string? fontNameAscii = null,
        [Description("Font name for Far East characters (Chinese/Japanese/Korean, optional)")]
        string? fontNameFarEast = null,
        [Description("Font size in points (optional)")]
        double? fontSize = null,
        [Description("Section index (0-based, optional, default: 0, use -1 to apply to all sections)")]
        int sectionIndex = 0,
        [Description(
            "Header/Footer type: primary (default), firstPage, evenPages. Use firstPage for different first page, evenPages for odd/even page layouts.")]
        string headerFooterType = "primary",
        [Description(
            "Make image floating instead of inline (optional, default: false, for image operations). Floating images can be precisely positioned.")]
        bool isFloating = false,
        [Description(
            "Automatically add center and right tab stops when using left/center/right text (optional, default: true, for text operations)")]
        bool autoTabStops = true,
        [Description("Clear existing content before setting new content (optional, default: true)")]
        bool clearExisting = true,
        [Description(
            "Only clear text content, preserve images and shapes (optional, default: false, for text operations)")]
        bool clearTextOnly = false,
        [Description("Remove existing images before adding new one (optional, default: true, for image operations)")]
        bool removeExisting = true)
    {
        var parameters = BuildParameters(operation, headerLeft, headerCenter, headerRight, footerLeft, footerCenter,
            footerRight, imagePath, alignment, imageWidth, imageHeight, lineStyle, lineWidth, tabStops, fontName,
            fontNameAscii, fontNameFarEast, fontSize, sectionIndex, headerFooterType, isFloating, autoTabStops,
            clearExisting, clearTextOnly, removeExisting);

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
        string? headerLeft,
        string? headerCenter,
        string? headerRight,
        string? footerLeft,
        string? footerCenter,
        string? footerRight,
        string? imagePath,
        string alignment,
        double? imageWidth,
        double? imageHeight,
        string lineStyle,
        double? lineWidth,
        JsonArray? tabStops,
        string? fontName,
        string? fontNameAscii,
        string? fontNameFarEast,
        double? fontSize,
        int sectionIndex,
        string headerFooterType,
        bool isFloating,
        bool autoTabStops,
        bool clearExisting,
        bool clearTextOnly,
        bool removeExisting)
    {
        var parameters = new OperationParameters();

        // Common parameters for all operations
        parameters.Set("sectionIndex", sectionIndex);
        parameters.Set("headerFooterType", headerFooterType);

        switch (operation.ToLower())
        {
            case "set_header_text":
                if (headerLeft != null) parameters.Set("headerLeft", headerLeft);
                if (headerCenter != null) parameters.Set("headerCenter", headerCenter);
                if (headerRight != null) parameters.Set("headerRight", headerRight);
                if (fontName != null) parameters.Set("fontName", fontName);
                if (fontNameAscii != null) parameters.Set("fontNameAscii", fontNameAscii);
                if (fontNameFarEast != null) parameters.Set("fontNameFarEast", fontNameFarEast);
                if (fontSize.HasValue) parameters.Set("fontSize", fontSize.Value);
                parameters.Set("autoTabStops", autoTabStops);
                parameters.Set("clearExisting", clearExisting);
                parameters.Set("clearTextOnly", clearTextOnly);
                break;

            case "set_footer_text":
                if (footerLeft != null) parameters.Set("footerLeft", footerLeft);
                if (footerCenter != null) parameters.Set("footerCenter", footerCenter);
                if (footerRight != null) parameters.Set("footerRight", footerRight);
                if (fontName != null) parameters.Set("fontName", fontName);
                if (fontNameAscii != null) parameters.Set("fontNameAscii", fontNameAscii);
                if (fontNameFarEast != null) parameters.Set("fontNameFarEast", fontNameFarEast);
                if (fontSize.HasValue) parameters.Set("fontSize", fontSize.Value);
                parameters.Set("autoTabStops", autoTabStops);
                parameters.Set("clearExisting", clearExisting);
                parameters.Set("clearTextOnly", clearTextOnly);
                break;

            case "set_header_image":
            case "set_footer_image":
                if (imagePath != null) parameters.Set("imagePath", imagePath);
                parameters.Set("alignment", alignment);
                if (imageWidth.HasValue) parameters.Set("imageWidth", imageWidth.Value);
                if (imageHeight.HasValue) parameters.Set("imageHeight", imageHeight.Value);
                parameters.Set("isFloating", isFloating);
                parameters.Set("removeExisting", removeExisting);
                break;

            case "set_header_line":
            case "set_footer_line":
                parameters.Set("lineStyle", lineStyle);
                if (lineWidth.HasValue) parameters.Set("lineWidth", lineWidth.Value);
                break;

            case "set_header_tabs":
            case "set_footer_tabs":
                if (tabStops != null) parameters.Set("tabStops", tabStops);
                break;

            case "set_header_footer":
                if (headerLeft != null) parameters.Set("headerLeft", headerLeft);
                if (headerCenter != null) parameters.Set("headerCenter", headerCenter);
                if (headerRight != null) parameters.Set("headerRight", headerRight);
                if (footerLeft != null) parameters.Set("footerLeft", footerLeft);
                if (footerCenter != null) parameters.Set("footerCenter", footerCenter);
                if (footerRight != null) parameters.Set("footerRight", footerRight);
                if (fontName != null) parameters.Set("fontName", fontName);
                if (fontNameAscii != null) parameters.Set("fontNameAscii", fontNameAscii);
                if (fontNameFarEast != null) parameters.Set("fontNameFarEast", fontNameFarEast);
                if (fontSize.HasValue) parameters.Set("fontSize", fontSize.Value);
                parameters.Set("autoTabStops", autoTabStops);
                parameters.Set("clearExisting", clearExisting);
                parameters.Set("clearTextOnly", clearTextOnly);
                break;

            case "get":
                // Only sectionIndex needed, already set above
                break;
        }

        return parameters;
    }
}
