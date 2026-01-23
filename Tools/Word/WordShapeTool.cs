using System.ComponentModel;
using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Tool for managing shapes (lines, textboxes, charts, etc.) in Word documents
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Word.Shape")]
[McpServerToolType]
public class WordShapeTool
{
    /// <summary>
    ///     Handler registry for shape operations
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
    ///     Initializes a new instance of the WordShapeTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation</param>
    public WordShapeTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Document>.CreateFromNamespace("AsposeMcpServer.Handlers.Word.Shape");
    }

    /// <summary>
    ///     Executes a Word shape operation (add_line, add_textbox, get_textboxes, edit_textbox_content, set_textbox_border,
    ///     add_chart, add, get, delete).
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: add_line, add_textbox, get_textboxes, edit_textbox_content,
    ///     set_textbox_border, add_chart, add, get, delete.
    /// </param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="location">Location: body, header, footer (for add_line, default: body).</param>
    /// <param name="position">Position: start, end (for add_line, default: end).</param>
    /// <param name="lineStyle">Line style: border, shape (for add_line, default: shape).</param>
    /// <param name="lineWidth">Line width in points (for add_line, default: 1.0).</param>
    /// <param name="lineColor">Line color hex (for add_line, default: 000000).</param>
    /// <param name="width">Width/length in points (for add_line: line length; for add: shape width).</param>
    /// <param name="text">Text content (for add_textbox, edit_textbox_content).</param>
    /// <param name="textboxWidth">Textbox width in points (for add_textbox, default: 200).</param>
    /// <param name="textboxHeight">Textbox height in points (for add_textbox, default: 100).</param>
    /// <param name="positionX">Horizontal position in points (for add_textbox, default: 100).</param>
    /// <param name="positionY">Vertical position in points (for add_textbox, default: 100).</param>
    /// <param name="backgroundColor">Background color hex (for add_textbox).</param>
    /// <param name="borderColor">Border color hex (for add_textbox, set_textbox_border).</param>
    /// <param name="borderWidth">Border width in points (for add_textbox, set_textbox_border, default: 1).</param>
    /// <param name="fontName">Font name.</param>
    /// <param name="fontSize">Font size in points.</param>
    /// <param name="fontNameAscii">Font name for ASCII characters.</param>
    /// <param name="fontNameFarEast">Font name for Far East characters.</param>
    /// <param name="bold">Bold text.</param>
    /// <param name="italic">Italic text.</param>
    /// <param name="color">Text color hex.</param>
    /// <param name="textAlignment">Text alignment: left, center, right.</param>
    /// <param name="textboxIndex">Textbox index (0-based, for textbox operations).</param>
    /// <param name="shapeIndex">Shape index (0-based, for general shape operations).</param>
    /// <param name="shapeType">Shape type: Rectangle, RoundedRectangle, Ellipse, etc. (for add).</param>
    /// <param name="height">Shape height in points (for add).</param>
    /// <param name="x">Shape X position in points (for add).</param>
    /// <param name="y">Shape Y position in points (for add).</param>
    /// <param name="appendText">Append text to existing content (for edit_textbox_content).</param>
    /// <param name="clearFormatting">Clear existing formatting (for edit_textbox_content).</param>
    /// <param name="borderVisible">Show border (for set_textbox_border).</param>
    /// <param name="borderStyle">Border style (for set_textbox_border).</param>
    /// <param name="includeContent">Include textbox content (for get_textboxes).</param>
    /// <param name="chartType">Chart type: Column, Bar, Line, Pie, etc. (for add_chart).</param>
    /// <param name="data">Chart data as JSON 2D array (for add_chart).</param>
    /// <param name="chartTitle">Chart title (for add_chart).</param>
    /// <param name="chartWidth">Chart width in points (for add_chart).</param>
    /// <param name="chartHeight">Chart height in points (for add_chart).</param>
    /// <param name="paragraphIndex">Paragraph index to insert after (for add_chart).</param>
    /// <param name="alignment">Chart alignment (for add_chart).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "word_shape",
        Title = "Word Shape Operations",
        Destructive = true,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(
        @"Manage shapes in Word documents. Supports 9 operations: add_line, add_textbox, get_textboxes, edit_textbox_content, set_textbox_border, add_chart, add, get, delete.

Note: All position/size values are in points (1 point = 1/72 inch, 72 points = 1 inch).
Important: Textbox operations (get_textboxes, edit_textbox_content, set_textbox_border) use a separate textbox-only index system.
General shape operations (add, get, delete) use an index that includes ALL shapes (lines, rectangles, textboxes, images, etc.).

Usage examples:
- Add line: word_shape(operation='add_line', path='doc.docx')
- Add textbox: word_shape(operation='add_textbox', path='doc.docx', text='Textbox content', positionX=100, positionY=100, textboxWidth=200, textboxHeight=100)
- Get textboxes: word_shape(operation='get_textboxes', path='doc.docx')
- Edit textbox: word_shape(operation='edit_textbox_content', path='doc.docx', textboxIndex=0, text='Updated content')
- Set border: word_shape(operation='set_textbox_border', path='doc.docx', textboxIndex=0, borderColor='#FF0000', borderWidth=2)
- Add chart: word_shape(operation='add_chart', path='doc.docx', chartType='Column', data=[['A','B'],['1','2']])
- Add generic shape: word_shape(operation='add', path='doc.docx', shapeType='Rectangle', width=100, height=50)
- Get all shapes: word_shape(operation='get', path='doc.docx')
- Delete shape: word_shape(operation='delete', path='doc.docx', shapeIndex=0)")]
    public object Execute(
        [Description(
            "Operation: add_line, add_textbox, get_textboxes, edit_textbox_content, set_textbox_border, add_chart, add, get, delete")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Location: body, header, footer (for add_line, default: body)")]
        string location = "body",
        [Description("Position: start, end (for add_line, default: end)")]
        string position = "end",
        [Description("Line style: border, shape (for add_line, default: shape)")]
        string lineStyle = "shape",
        [Description("Line width in points (for add_line, default: 1.0)")]
        double lineWidth = 1.0,
        [Description("Line color hex (for add_line, default: 000000)")]
        string lineColor = "000000",
        [Description("Width/length in points (for add_line: line length; for add: shape width)")]
        double? width = null,
        [Description("Text content (for add_textbox, edit_textbox_content)")]
        string? text = null,
        [Description("Textbox width in points (for add_textbox, default: 200)")]
        double textboxWidth = 200,
        [Description("Textbox height in points (for add_textbox, default: 100)")]
        double textboxHeight = 100,
        [Description("Horizontal position in points (for add_textbox, default: 100)")]
        double positionX = 100,
        [Description("Vertical position in points (for add_textbox, default: 100)")]
        double positionY = 100,
        [Description("Background color hex (for add_textbox)")]
        string? backgroundColor = null,
        [Description("Border color hex (for add_textbox, set_textbox_border)")]
        string? borderColor = null,
        [Description("Border width in points (for add_textbox, set_textbox_border, default: 1)")]
        double borderWidth = 1,
        [Description("Font name (for add_textbox, edit_textbox_content)")]
        string? fontName = null,
        [Description("Font name for ASCII characters (for add_textbox, edit_textbox_content)")]
        string? fontNameAscii = null,
        [Description("Font name for Far East characters (for add_textbox, edit_textbox_content)")]
        string? fontNameFarEast = null,
        [Description("Font size in points (for add_textbox, edit_textbox_content)")]
        double? fontSize = null,
        [Description("Bold text (for add_textbox, edit_textbox_content)")]
        bool? bold = null,
        [Description("Italic text (for edit_textbox_content)")]
        bool? italic = null,
        [Description("Text color hex (for edit_textbox_content)")]
        string? color = null,
        [Description("Text alignment: left, center, right (for add_textbox, default: left)")]
        string textAlignment = "left",
        [Description("Textbox index (0-based, textbox-only index for edit_textbox_content, set_textbox_border)")]
        int? textboxIndex = null,
        [Description("Shape index (0-based, global index including all shapes, for delete operation)")]
        int? shapeIndex = null,
        [Description("Shape type: rectangle, ellipse, roundrectangle, line (for add operation)")]
        string? shapeType = null,
        [Description("Shape height in points (for add operation)")]
        double? height = null,
        [Description("Shape X position in points (for add operation, default: 100)")]
        double x = 100,
        [Description("Shape Y position in points (for add operation, default: 100)")]
        double y = 100,
        [Description("Append text to existing content (for edit_textbox_content, default: false)")]
        bool appendText = false,
        [Description("Clear existing formatting (for edit_textbox_content, default: false)")]
        bool clearFormatting = false,
        [Description("Show border (for set_textbox_border, default: true)")]
        bool borderVisible = true,
        [Description(
            "Border style: solid, dash, dot, dashDot, dashDotDot, roundDot (for set_textbox_border, default: solid)")]
        string borderStyle = "solid",
        [Description("Include textbox content (for get_textboxes, default: true)")]
        bool includeContent = true,
        [Description("Chart type: column, bar, line, pie, area, scatter, doughnut (for add_chart, default: column)")]
        string chartType = "column",
        [Description("Chart data as 2D array (for add_chart)")]
        string[][]? data = null,
        [Description("Chart title (for add_chart, optional)")]
        string? chartTitle = null,
        [Description("Chart width in points (for add_chart, default: 432)")]
        double chartWidth = 432,
        [Description("Chart height in points (for add_chart, default: 252)")]
        double chartHeight = 252,
        [Description("Paragraph index to insert after (for add_chart, optional, use -1 for beginning)")]
        int? paragraphIndex = null,
        [Description("Chart alignment: left, center, right (for add_chart, default: left)")]
        string alignment = "left")
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, location, position, lineStyle, lineWidth, lineColor, width,
            text, textboxWidth, textboxHeight, positionX, positionY, backgroundColor, borderColor, borderWidth,
            fontName, fontNameAscii, fontNameFarEast, fontSize, bold, italic, color, textAlignment, textboxIndex,
            shapeIndex, shapeType, height, x, y, appendText, clearFormatting, borderVisible, borderStyle,
            includeContent, chartType, data, chartTitle, chartWidth, chartHeight, paragraphIndex, alignment);

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
        string location,
        string position,
        string lineStyle,
        double lineWidth,
        string lineColor,
        double? width,
        string? text,
        double textboxWidth,
        double textboxHeight,
        double positionX,
        double positionY,
        string? backgroundColor,
        string? borderColor,
        double borderWidth,
        string? fontName,
        string? fontNameAscii,
        string? fontNameFarEast,
        double? fontSize,
        bool? bold,
        bool? italic,
        string? color,
        string textAlignment,
        int? textboxIndex,
        int? shapeIndex,
        string? shapeType,
        double? height,
        double x,
        double y,
        bool appendText,
        bool clearFormatting,
        bool borderVisible,
        string borderStyle,
        bool includeContent,
        string chartType,
        string[][]? data,
        string? chartTitle,
        double chartWidth,
        double chartHeight,
        int? paragraphIndex,
        string alignment)
    {
        var parameters = new OperationParameters();

        return operation.ToLower() switch
        {
            "add_line" => BuildAddLineParameters(parameters, location, position, lineStyle, lineWidth, lineColor,
                width),
            "add_textbox" => BuildAddTextboxParameters(parameters, text, textboxWidth, textboxHeight, positionX,
                positionY, backgroundColor, borderColor, borderWidth, fontName, fontNameAscii, fontNameFarEast,
                fontSize,
                bold, textAlignment),
            "get_textboxes" => BuildGetTextboxesParameters(parameters, includeContent),
            "edit_textbox_content" => BuildEditTextboxContentParameters(parameters, textboxIndex, text, appendText,
                fontName, fontNameAscii, fontNameFarEast, fontSize, bold, italic, color, clearFormatting),
            "set_textbox_border" => BuildSetTextboxBorderParameters(parameters, textboxIndex, borderVisible,
                borderColor,
                borderWidth, borderStyle),
            "add_chart" => BuildAddChartParameters(parameters, chartType, data, chartTitle, chartWidth, chartHeight,
                paragraphIndex, alignment),
            "add" => BuildAddShapeParameters(parameters, shapeType, width, height, x, y),
            "get" => parameters,
            "delete" => BuildDeleteParameters(parameters, shapeIndex),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the add line operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="location">The location: 'body', 'header', 'footer'.</param>
    /// <param name="position">The position: 'start', 'end'.</param>
    /// <param name="lineStyle">The line style: 'border', 'shape'.</param>
    /// <param name="lineWidth">The line width in points.</param>
    /// <param name="lineColor">The line color in hex format.</param>
    /// <param name="width">The line length in points.</param>
    /// <returns>OperationParameters configured for the add line operation.</returns>
    private static OperationParameters BuildAddLineParameters(OperationParameters parameters, string location,
        string position, string lineStyle, double lineWidth, string lineColor, double? width)
    {
        parameters.Set("location", location);
        parameters.Set("position", position);
        parameters.Set("lineStyle", lineStyle);
        parameters.Set("lineWidth", lineWidth);
        parameters.Set("lineColor", lineColor);
        if (width.HasValue) parameters.Set("width", width.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the add textbox operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="text">The text content.</param>
    /// <param name="textboxWidth">The textbox width in points.</param>
    /// <param name="textboxHeight">The textbox height in points.</param>
    /// <param name="positionX">The horizontal position in points.</param>
    /// <param name="positionY">The vertical position in points.</param>
    /// <param name="backgroundColor">The background color in hex format.</param>
    /// <param name="borderColor">The border color in hex format.</param>
    /// <param name="borderWidth">The border width in points.</param>
    /// <param name="fontName">The font name.</param>
    /// <param name="fontNameAscii">The font name for ASCII characters.</param>
    /// <param name="fontNameFarEast">The font name for Far East characters.</param>
    /// <param name="fontSize">The font size in points.</param>
    /// <param name="bold">Whether the text is bold.</param>
    /// <param name="textAlignment">The text alignment: 'left', 'center', 'right'.</param>
    /// <returns>OperationParameters configured for the add textbox operation.</returns>
    private static OperationParameters BuildAddTextboxParameters(
        OperationParameters parameters, string? text,
        double textboxWidth, double textboxHeight, double positionX, double positionY, string? backgroundColor,
        string? borderColor, double borderWidth, string? fontName, string? fontNameAscii, string? fontNameFarEast,
        double? fontSize, bool? bold, string textAlignment)
    {
        if (text != null) parameters.Set("text", text);
        parameters.Set("textboxWidth", textboxWidth);
        parameters.Set("textboxHeight", textboxHeight);
        parameters.Set("positionX", positionX);
        parameters.Set("positionY", positionY);
        if (backgroundColor != null) parameters.Set("backgroundColor", backgroundColor);
        if (borderColor != null) parameters.Set("borderColor", borderColor);
        parameters.Set("borderWidth", borderWidth);
        if (fontName != null) parameters.Set("fontName", fontName);
        if (fontNameAscii != null) parameters.Set("fontNameAscii", fontNameAscii);
        if (fontNameFarEast != null) parameters.Set("fontNameFarEast", fontNameFarEast);
        if (fontSize.HasValue) parameters.Set("fontSize", fontSize.Value);
        if (bold.HasValue) parameters.Set("bold", bold.Value);
        parameters.Set("textAlignment", textAlignment);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the get textboxes operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="includeContent">Whether to include textbox content in the results.</param>
    /// <returns>OperationParameters configured for the get textboxes operation.</returns>
    private static OperationParameters BuildGetTextboxesParameters(OperationParameters parameters, bool includeContent)
    {
        parameters.Set("includeContent", includeContent);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the edit textbox content operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="textboxIndex">The textbox index (0-based).</param>
    /// <param name="text">The text content.</param>
    /// <param name="appendText">Whether to append text to existing content.</param>
    /// <param name="fontName">The font name.</param>
    /// <param name="fontNameAscii">The font name for ASCII characters.</param>
    /// <param name="fontNameFarEast">The font name for Far East characters.</param>
    /// <param name="fontSize">The font size in points.</param>
    /// <param name="bold">Whether the text is bold.</param>
    /// <param name="italic">Whether the text is italic.</param>
    /// <param name="color">The text color in hex format.</param>
    /// <param name="clearFormatting">Whether to clear existing formatting.</param>
    /// <returns>OperationParameters configured for the edit textbox content operation.</returns>
    private static OperationParameters
        BuildEditTextboxContentParameters(
            OperationParameters parameters,
            int? textboxIndex, string? text, bool appendText, string? fontName, string? fontNameAscii,
            string? fontNameFarEast, double? fontSize, bool? bold, bool? italic, string? color, bool clearFormatting)
    {
        if (textboxIndex.HasValue) parameters.Set("textboxIndex", textboxIndex.Value);
        if (text != null) parameters.Set("text", text);
        parameters.Set("appendText", appendText);
        if (fontName != null) parameters.Set("fontName", fontName);
        if (fontNameAscii != null) parameters.Set("fontNameAscii", fontNameAscii);
        if (fontNameFarEast != null) parameters.Set("fontNameFarEast", fontNameFarEast);
        if (fontSize.HasValue) parameters.Set("fontSize", fontSize.Value);
        if (bold.HasValue) parameters.Set("bold", bold.Value);
        if (italic.HasValue) parameters.Set("italic", italic.Value);
        if (color != null) parameters.Set("color", color);
        parameters.Set("clearFormatting", clearFormatting);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the set textbox border operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="textboxIndex">The textbox index (0-based).</param>
    /// <param name="borderVisible">Whether the border is visible.</param>
    /// <param name="borderColor">The border color in hex format.</param>
    /// <param name="borderWidth">The border width in points.</param>
    /// <param name="borderStyle">The border style: 'solid', 'dash', 'dot', etc.</param>
    /// <returns>OperationParameters configured for the set textbox border operation.</returns>
    private static OperationParameters BuildSetTextboxBorderParameters(OperationParameters parameters,
        int? textboxIndex, bool borderVisible, string? borderColor, double borderWidth, string borderStyle)
    {
        if (textboxIndex.HasValue) parameters.Set("textboxIndex", textboxIndex.Value);
        parameters.Set("borderVisible", borderVisible);
        if (borderColor != null) parameters.Set("borderColor", borderColor);
        parameters.Set("borderWidth", borderWidth);
        parameters.Set("borderStyle", borderStyle);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the add chart operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="chartType">The chart type: 'column', 'bar', 'line', 'pie', etc.</param>
    /// <param name="data">The chart data as 2D array.</param>
    /// <param name="chartTitle">The chart title.</param>
    /// <param name="chartWidth">The chart width in points.</param>
    /// <param name="chartHeight">The chart height in points.</param>
    /// <param name="paragraphIndex">The paragraph index to insert after.</param>
    /// <param name="alignment">The chart alignment: 'left', 'center', 'right'.</param>
    /// <returns>OperationParameters configured for the add chart operation.</returns>
    private static OperationParameters BuildAddChartParameters(
        OperationParameters parameters, string chartType,
        string[][]? data, string? chartTitle, double chartWidth, double chartHeight, int? paragraphIndex,
        string alignment)
    {
        parameters.Set("chartType", chartType);
        if (data != null) parameters.Set("data", data);
        if (chartTitle != null) parameters.Set("chartTitle", chartTitle);
        parameters.Set("chartWidth", chartWidth);
        parameters.Set("chartHeight", chartHeight);
        if (paragraphIndex.HasValue) parameters.Set("paragraphIndex", paragraphIndex.Value);
        parameters.Set("alignment", alignment);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the add shape operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="shapeType">The shape type: 'rectangle', 'ellipse', 'roundrectangle', 'line'.</param>
    /// <param name="width">The shape width in points.</param>
    /// <param name="height">The shape height in points.</param>
    /// <param name="x">The shape X position in points.</param>
    /// <param name="y">The shape Y position in points.</param>
    /// <returns>OperationParameters configured for the add shape operation.</returns>
    private static OperationParameters BuildAddShapeParameters(OperationParameters parameters, string? shapeType,
        double? width, double? height, double x, double y)
    {
        if (shapeType != null) parameters.Set("shapeType", shapeType);
        if (width.HasValue) parameters.Set("width", width.Value);
        if (height.HasValue) parameters.Set("height", height.Value);
        parameters.Set("x", x);
        parameters.Set("y", y);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete shape operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters.</param>
    /// <param name="shapeIndex">The shape index (0-based, global index including all shapes).</param>
    /// <returns>OperationParameters configured for the delete operation.</returns>
    private static OperationParameters BuildDeleteParameters(OperationParameters parameters, int? shapeIndex)
    {
        if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
        return parameters;
    }
}
