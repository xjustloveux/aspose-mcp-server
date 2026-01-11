using System.ComponentModel;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Handlers.Word.Format;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for formatting text and paragraphs in Word documents
/// </summary>
[McpServerToolType]
public class WordFormatTool
{
    /// <summary>
    ///     Handler registry for format operations
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
    ///     Initializes a new instance of the WordFormatTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordFormatTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = WordFormatHandlerRegistry.Create();
    }

    /// <summary>
    ///     Executes a Word format operation (get_run_format, set_run_format, get_tab_stops, add_tab_stop, clear_tab_stops,
    ///     set_paragraph_border).
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: get_run_format, set_run_format, get_tab_stops, add_tab_stop,
    ///     clear_tab_stops, set_paragraph_border.
    /// </param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="paragraphIndex">Paragraph index (0-based).</param>
    /// <param name="runIndex">Run index within paragraph (0-based, optional).</param>
    /// <param name="sectionIndex">Section index (0-based, default: 0).</param>
    /// <param name="includeInherited">Include inherited format from paragraph/style (for get_run_format, default: false).</param>
    /// <param name="fontName">Font name (for set_run_format).</param>
    /// <param name="fontNameAscii">Font name for ASCII characters (for set_run_format).</param>
    /// <param name="fontNameFarEast">Font name for Far East characters (for set_run_format).</param>
    /// <param name="fontSize">Font size in points (for set_run_format).</param>
    /// <param name="bold">Bold text (for set_run_format).</param>
    /// <param name="italic">Italic text (for set_run_format).</param>
    /// <param name="underline">Underline style (for set_run_format).</param>
    /// <param name="color">Text color hex (for set_run_format).</param>
    /// <param name="tabPosition">Tab stop position in points (for add_tab_stop).</param>
    /// <param name="tabAlignment">Tab alignment: left, center, right, decimal (for add_tab_stop).</param>
    /// <param name="tabLeader">Tab leader: none, dots, dashes, line (for add_tab_stop).</param>
    /// <param name="borderPosition">Border position: all, top, bottom, left, right (for set_paragraph_border).</param>
    /// <param name="borderTop">Show top border (for set_paragraph_border).</param>
    /// <param name="borderBottom">Show bottom border (for set_paragraph_border).</param>
    /// <param name="borderLeft">Show left border (for set_paragraph_border).</param>
    /// <param name="borderRight">Show right border (for set_paragraph_border).</param>
    /// <param name="lineStyle">Border line style: single, double, thick (for set_paragraph_border).</param>
    /// <param name="lineWidth">Border line width in points (for set_paragraph_border).</param>
    /// <param name="lineColor">Border color hex (for set_paragraph_border).</param>
    /// <param name="location">Where to get tab stops from: header, footer, body (for get_tab_stops).</param>
    /// <param name="allParagraphs">Read tab stops from all paragraphs (for get_tab_stops).</param>
    /// <param name="includeStyle">Include tab stops from paragraph style (for get_tab_stops).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "word_format")]
    [Description(
        @"Manage formatting in Word documents. Supports 6 operations: get_run_format, set_run_format, get_tab_stops, add_tab_stop, clear_tab_stops, set_paragraph_border.

Usage examples:
- Get run format: word_format(operation='get_run_format', path='doc.docx', paragraphIndex=0, runIndex=0)
- Get inherited format: word_format(operation='get_run_format', path='doc.docx', paragraphIndex=0, runIndex=0, includeInherited=true)
- Set run format: word_format(operation='set_run_format', path='doc.docx', paragraphIndex=0, runIndex=0, bold=true, fontSize=14)
- Reset color to auto: word_format(operation='set_run_format', path='doc.docx', paragraphIndex=0, runIndex=0, color='auto')
- Get tab stops: word_format(operation='get_tab_stops', path='doc.docx', paragraphIndex=0)
- Add tab stop: word_format(operation='add_tab_stop', path='doc.docx', paragraphIndex=0, tabPosition=72, tabAlignment='center')
- Clear tab stops: word_format(operation='clear_tab_stops', path='doc.docx', paragraphIndex=0)
- Set paragraph border: word_format(operation='set_paragraph_border', path='doc.docx', paragraphIndex=0, borderPosition='all', lineStyle='single', lineWidth=1.0)")]
    public string Execute(
        [Description(@"Operation to perform.
- 'get_run_format': Get run formatting (required params: path, paragraphIndex, runIndex)
- 'set_run_format': Set run formatting (required params: path, paragraphIndex, runIndex)
- 'get_tab_stops': Get tab stops (required params: path, paragraphIndex)
- 'add_tab_stop': Add a tab stop (required params: path, paragraphIndex, tabPosition)
- 'clear_tab_stops': Clear tab stops (required params: path, paragraphIndex)
- 'set_paragraph_border': Set paragraph border (required params: path, paragraphIndex)")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Paragraph index (0-based)")]
        int? paragraphIndex = null,
        [Description("Run index within paragraph (0-based, optional)")]
        int? runIndex = null,
        [Description("Section index (0-based, default: 0)")]
        int sectionIndex = 0,
        [Description(
            "Include inherited format from paragraph/style (for get_run_format, default: false). When true, shows the effective computed format.")]
        bool includeInherited = false,
        [Description("Font name (for set_run_format)")]
        string? fontName = null,
        [Description("Font name for ASCII characters (for set_run_format)")]
        string? fontNameAscii = null,
        [Description("Font name for Far East characters (for set_run_format)")]
        string? fontNameFarEast = null,
        [Description("Font size in points (for set_run_format)")]
        double? fontSize = null,
        [Description("Bold text (for set_run_format)")]
        bool? bold = null,
        [Description("Italic text (for set_run_format)")]
        bool? italic = null,
        [Description("Underline text (for set_run_format)")]
        bool? underline = null,
        [Description("Font color hex (for set_run_format)")]
        string? color = null,
        [Description("Where to get tab stops from: header, footer, body (for get_tab_stops, default: body)")]
        string location = "body",
        [Description("Read tab stops from all paragraphs (for get_tab_stops, default: false)")]
        bool allParagraphs = false,
        [Description("Include tab stops from paragraph style (for get_tab_stops, default: true)")]
        bool includeStyle = true,
        [Description("Tab stop position in points (for add_tab_stop, required)")]
        double? tabPosition = null,
        [Description("Tab stop alignment (for add_tab_stop, default: left)")]
        string tabAlignment = "left",
        [Description("Tab stop leader character (for add_tab_stop, default: none)")]
        string tabLeader = "none",
        [Description(
            "Border position shortcut (for set_paragraph_border): 'all', 'top-bottom', 'left-right', 'box'. Overrides individual border flags.")]
        string? borderPosition = null,
        [Description("Show top border (for set_paragraph_border, default: false)")]
        bool borderTop = false,
        [Description("Show bottom border (for set_paragraph_border, default: false)")]
        bool borderBottom = false,
        [Description("Show left border (for set_paragraph_border, default: false)")]
        bool borderLeft = false,
        [Description("Show right border (for set_paragraph_border, default: false)")]
        bool borderRight = false,
        [Description("Border line style: none, single, double, dotted, dashed, thick (for set_paragraph_border)")]
        string lineStyle = "single",
        [Description("Border line width in points (for set_paragraph_border, default: 0.5)")]
        double lineWidth = 0.5,
        [Description("Border line color hex (for set_paragraph_border, default: 000000)")]
        string lineColor = "000000")
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, paragraphIndex, runIndex, sectionIndex, includeInherited,
            fontName, fontNameAscii, fontNameFarEast, fontSize, bold, italic, underline, color,
            location, allParagraphs, includeStyle, tabPosition, tabAlignment, tabLeader,
            borderPosition, borderTop, borderBottom, borderLeft, borderRight, lineStyle, lineWidth, lineColor);

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
        if (operation.ToLower() is "get_run_format" or "get_tab_stops")
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
        int? runIndex,
        int sectionIndex,
        bool includeInherited,
        string? fontName,
        string? fontNameAscii,
        string? fontNameFarEast,
        double? fontSize,
        bool? bold,
        bool? italic,
        bool? underline,
        string? color,
        string location,
        bool allParagraphs,
        bool includeStyle,
        double? tabPosition,
        string tabAlignment,
        string tabLeader,
        string? borderPosition,
        bool borderTop,
        bool borderBottom,
        bool borderLeft,
        bool borderRight,
        string lineStyle,
        double lineWidth,
        string lineColor)
    {
        var parameters = new OperationParameters();

        switch (operation.ToLower())
        {
            case "get_run_format":
                parameters.Set("paragraphIndex", paragraphIndex ?? 0);
                if (runIndex.HasValue) parameters.Set("runIndex", runIndex.Value);
                parameters.Set("includeInherited", includeInherited);
                break;

            case "set_run_format":
                parameters.Set("paragraphIndex", paragraphIndex ?? 0);
                if (runIndex.HasValue) parameters.Set("runIndex", runIndex.Value);
                if (fontName != null) parameters.Set("fontName", fontName);
                if (fontNameAscii != null) parameters.Set("fontNameAscii", fontNameAscii);
                if (fontNameFarEast != null) parameters.Set("fontNameFarEast", fontNameFarEast);
                if (fontSize.HasValue) parameters.Set("fontSize", fontSize.Value);
                if (bold.HasValue) parameters.Set("bold", bold.Value);
                if (italic.HasValue) parameters.Set("italic", italic.Value);
                if (underline.HasValue) parameters.Set("underline", underline.Value);
                if (color != null) parameters.Set("color", color);
                break;

            case "get_tab_stops":
                parameters.Set("location", location);
                parameters.Set("paragraphIndex", paragraphIndex ?? 0);
                parameters.Set("sectionIndex", sectionIndex);
                parameters.Set("allParagraphs", allParagraphs);
                parameters.Set("includeStyle", includeStyle);
                break;

            case "add_tab_stop":
                parameters.Set("paragraphIndex", paragraphIndex ?? 0);
                parameters.Set("tabPosition", tabPosition ?? 0.0);
                parameters.Set("tabAlignment", tabAlignment);
                parameters.Set("tabLeader", tabLeader);
                break;

            case "clear_tab_stops":
                parameters.Set("paragraphIndex", paragraphIndex ?? 0);
                break;

            case "set_paragraph_border":
                parameters.Set("paragraphIndex", paragraphIndex ?? 0);
                if (borderPosition != null) parameters.Set("borderPosition", borderPosition);
                parameters.Set("borderTop", borderTop);
                parameters.Set("borderBottom", borderBottom);
                parameters.Set("borderLeft", borderLeft);
                parameters.Set("borderRight", borderRight);
                parameters.Set("lineStyle", lineStyle);
                parameters.Set("lineWidth", lineWidth);
                parameters.Set("lineColor", lineColor);
                break;
        }

        return parameters;
    }
}
