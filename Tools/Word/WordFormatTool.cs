using System.ComponentModel;
using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for formatting text and paragraphs in Word documents
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Word.Format")]
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
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Word.Format");
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
    [McpServerTool(
        Name = "word_format",
        Title = "Word Format Operations",
        Destructive = false,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
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
    public object Execute(
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
            return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
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
        return operation.ToLower() switch
        {
            "get_run_format" => BuildGetRunFormatParameters(paragraphIndex, runIndex, includeInherited),
            "set_run_format" => BuildSetRunFormatParameters(paragraphIndex, runIndex, fontName, fontNameAscii,
                fontNameFarEast, fontSize, bold, italic, underline, color),
            "get_tab_stops" => BuildGetTabStopsParameters(location, paragraphIndex, sectionIndex, allParagraphs,
                includeStyle),
            "add_tab_stop" => BuildAddTabStopParameters(paragraphIndex, tabPosition, tabAlignment, tabLeader),
            "clear_tab_stops" => BuildClearTabStopsParameters(paragraphIndex),
            "set_paragraph_border" => BuildSetParagraphBorderParameters(paragraphIndex, borderPosition, borderTop,
                borderBottom, borderLeft, borderRight, lineStyle, lineWidth, lineColor),
            _ => new OperationParameters()
        };
    }

    /// <summary>
    ///     Builds parameters for the get_run_format operation.
    /// </summary>
    /// <param name="paragraphIndex">The paragraph index (0-based).</param>
    /// <param name="runIndex">The run index within the paragraph (0-based).</param>
    /// <param name="includeInherited">Whether to include inherited format from paragraph/style.</param>
    /// <returns>OperationParameters configured for getting run format.</returns>
    private static OperationParameters BuildGetRunFormatParameters(int? paragraphIndex, int? runIndex,
        bool includeInherited)
    {
        var parameters = new OperationParameters();
        parameters.Set("paragraphIndex", paragraphIndex ?? 0);
        if (runIndex.HasValue) parameters.Set("runIndex", runIndex.Value);
        parameters.Set("includeInherited", includeInherited);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the set_run_format operation.
    /// </summary>
    /// <param name="paragraphIndex">The paragraph index (0-based).</param>
    /// <param name="runIndex">The run index within the paragraph (0-based).</param>
    /// <param name="fontName">The font name.</param>
    /// <param name="fontNameAscii">The font name for ASCII characters.</param>
    /// <param name="fontNameFarEast">The font name for Far East characters.</param>
    /// <param name="fontSize">The font size in points.</param>
    /// <param name="bold">Whether the text is bold.</param>
    /// <param name="italic">Whether the text is italic.</param>
    /// <param name="underline">Whether the text is underlined.</param>
    /// <param name="color">The text color in hex format.</param>
    /// <returns>OperationParameters configured for setting run format.</returns>
    private static OperationParameters BuildSetRunFormatParameters(
        int? paragraphIndex, int? runIndex, string? fontName,
        string? fontNameAscii, string? fontNameFarEast, double? fontSize, bool? bold, bool? italic, bool? underline,
        string? color)
    {
        var parameters = new OperationParameters();
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
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the get_tab_stops operation.
    /// </summary>
    /// <param name="location">Where to get tab stops from: header, footer, body.</param>
    /// <param name="paragraphIndex">The paragraph index (0-based).</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <param name="allParagraphs">Whether to read tab stops from all paragraphs.</param>
    /// <param name="includeStyle">Whether to include tab stops from paragraph style.</param>
    /// <returns>OperationParameters configured for getting tab stops.</returns>
    private static OperationParameters BuildGetTabStopsParameters(string location, int? paragraphIndex,
        int sectionIndex,
        bool allParagraphs, bool includeStyle)
    {
        var parameters = new OperationParameters();
        parameters.Set("location", location);
        parameters.Set("paragraphIndex", paragraphIndex ?? 0);
        parameters.Set("sectionIndex", sectionIndex);
        parameters.Set("allParagraphs", allParagraphs);
        parameters.Set("includeStyle", includeStyle);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the add_tab_stop operation.
    /// </summary>
    /// <param name="paragraphIndex">The paragraph index (0-based).</param>
    /// <param name="tabPosition">The tab stop position in points.</param>
    /// <param name="tabAlignment">The tab alignment: left, center, right, decimal.</param>
    /// <param name="tabLeader">The tab leader: none, dots, dashes, line.</param>
    /// <returns>OperationParameters configured for adding a tab stop.</returns>
    private static OperationParameters BuildAddTabStopParameters(int? paragraphIndex, double? tabPosition,
        string tabAlignment, string tabLeader)
    {
        var parameters = new OperationParameters();
        parameters.Set("paragraphIndex", paragraphIndex ?? 0);
        parameters.Set("tabPosition", tabPosition ?? 0.0);
        parameters.Set("tabAlignment", tabAlignment);
        parameters.Set("tabLeader", tabLeader);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the clear_tab_stops operation.
    /// </summary>
    /// <param name="paragraphIndex">The paragraph index (0-based).</param>
    /// <returns>OperationParameters configured for clearing tab stops.</returns>
    private static OperationParameters BuildClearTabStopsParameters(int? paragraphIndex)
    {
        var parameters = new OperationParameters();
        parameters.Set("paragraphIndex", paragraphIndex ?? 0);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the set_paragraph_border operation.
    /// </summary>
    /// <param name="paragraphIndex">The paragraph index (0-based).</param>
    /// <param name="borderPosition">The border position shortcut: all, top-bottom, left-right, box.</param>
    /// <param name="borderTop">Whether to show the top border.</param>
    /// <param name="borderBottom">Whether to show the bottom border.</param>
    /// <param name="borderLeft">Whether to show the left border.</param>
    /// <param name="borderRight">Whether to show the right border.</param>
    /// <param name="lineStyle">The border line style: none, single, double, dotted, dashed, thick.</param>
    /// <param name="lineWidth">The border line width in points.</param>
    /// <param name="lineColor">The border line color in hex format.</param>
    /// <returns>OperationParameters configured for setting paragraph border.</returns>
    private static OperationParameters
        BuildSetParagraphBorderParameters(
            int? paragraphIndex, string? borderPosition,
            bool borderTop, bool borderBottom, bool borderLeft, bool borderRight, string lineStyle, double lineWidth,
            string lineColor)
    {
        var parameters = new OperationParameters();
        parameters.Set("paragraphIndex", paragraphIndex ?? 0);
        if (borderPosition != null) parameters.Set("borderPosition", borderPosition);
        parameters.Set("borderTop", borderTop);
        parameters.Set("borderBottom", borderBottom);
        parameters.Set("borderLeft", borderLeft);
        parameters.Set("borderRight", borderRight);
        parameters.Set("lineStyle", lineStyle);
        parameters.Set("lineWidth", lineWidth);
        parameters.Set("lineColor", lineColor);
        return parameters;
    }
}
