using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel sparklines (add, delete, get, set_style).
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Excel.Sparkline")]
[McpServerToolType]
public class ExcelSparklineTool
{
    /// <summary>
    ///     Handler registry for sparkline operations.
    /// </summary>
    private readonly HandlerRegistry<Workbook> _handlerRegistry;

    /// <summary>
    ///     Session identity accessor for session isolation support.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Document session manager for in-memory editing support.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelSparklineTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelSparklineTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.Sparkline");
    }

    /// <summary>
    ///     Executes an Excel sparkline operation (add, delete, get, set_style).
    /// </summary>
    /// <param name="operation">The operation to perform.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based).</param>
    /// <param name="dataRange">Data range for sparkline (for add, e.g., 'A1:A10').</param>
    /// <param name="locationRange">Location cell for sparkline (for add, e.g., 'B1').</param>
    /// <param name="type">Sparkline type: line, column, stacked/win_loss (for add, default: line).</param>
    /// <param name="groupIndex">Sparkline group index (for delete/set_style).</param>
    /// <param name="presetStyle">Preset style name (for set_style, e.g., 'Style1').</param>
    /// <param name="showHighPoint">Show high point marker (for set_style).</param>
    /// <param name="showLowPoint">Show low point marker (for set_style).</param>
    /// <param name="showFirstPoint">Show first point marker (for set_style).</param>
    /// <param name="showLastPoint">Show last point marker (for set_style).</param>
    /// <param name="showNegativePoints">Show negative point markers (for set_style).</param>
    /// <param name="isVertical">Whether data is arranged vertically (for add, auto-detected from data range if not specified).</param>
    /// <param name="showMarkers">Show data markers (for set_style).</param>
    /// <returns>A message or data indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "excel_sparkline",
        Title = "Excel Sparkline Operations",
        Destructive = true,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Manage Excel sparklines. Supports 4 operations: add, delete, get, set_style.

Usage examples:
- Add sparkline: excel_sparkline(operation='add', path='book.xlsx', dataRange='A1:A10', locationRange='B1')
- Add column sparkline: excel_sparkline(operation='add', path='book.xlsx', dataRange='A1:A10', locationRange='B1', type='column')
- Get sparklines: excel_sparkline(operation='get', path='book.xlsx')
- Delete sparkline: excel_sparkline(operation='delete', path='book.xlsx', groupIndex=0)
- Set style: excel_sparkline(operation='set_style', path='book.xlsx', groupIndex=0, showHighPoint=true)")]
    public object Execute(
        [Description(@"Operation to perform.
- 'add': Add a sparkline group (required params: dataRange, locationRange)
- 'get': Get sparkline groups information
- 'delete': Delete a sparkline group (required params: groupIndex)
- 'set_style': Set sparkline group style (required params: groupIndex)")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Data range for sparkline (for add, e.g., 'A1:A10')")]
        string? dataRange = null,
        [Description("Location cell for sparkline (for add, e.g., 'B1')")]
        string? locationRange = null,
        [Description("Sparkline type: line, column, stacked/win_loss (for add, default: line)")]
        string type = "line",
        [Description("Sparkline group index (for delete/set_style)")]
        int? groupIndex = null,
        [Description("Preset style name (for set_style, e.g., 'Style1')")]
        string? presetStyle = null,
        [Description("Show high point marker (for set_style)")]
        bool? showHighPoint = null,
        [Description("Show low point marker (for set_style)")]
        bool? showLowPoint = null,
        [Description("Show first point marker (for set_style)")]
        bool? showFirstPoint = null,
        [Description("Show last point marker (for set_style)")]
        bool? showLastPoint = null,
        [Description("Show negative point markers (for set_style)")]
        bool? showNegativePoints = null,
        [Description("Show data markers (for set_style)")]
        bool? showMarkers = null,
        [Description(
            "Whether data is arranged vertically/by column (for add, auto-detected from data range if not specified)")]
        bool? isVertical = null)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, sheetIndex, dataRange, locationRange, type, groupIndex,
            presetStyle, showHighPoint, showLowPoint, showFirstPoint, showLastPoint, showNegativePoints, showMarkers,
            isVertical);

        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Workbook>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);

        if (string.Equals(operation, "get", StringComparison.OrdinalIgnoreCase))
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
        string operation, int sheetIndex, string? dataRange, string? locationRange, string type,
        int? groupIndex, string? presetStyle, bool? showHighPoint, bool? showLowPoint,
        bool? showFirstPoint, bool? showLastPoint, bool? showNegativePoints, bool? showMarkers,
        bool? isVertical)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);

        return operation.ToLowerInvariant() switch
        {
            "add" => BuildAddParameters(parameters, dataRange, locationRange, type, isVertical),
            "delete" => BuildDeleteParameters(parameters, groupIndex),
            "get" => parameters,
            "set_style" => BuildSetStyleParameters(parameters, groupIndex, presetStyle, showHighPoint, showLowPoint,
                showFirstPoint, showLastPoint, showNegativePoints, showMarkers),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the add operation.
    /// </summary>
    private static OperationParameters BuildAddParameters(OperationParameters parameters, string? dataRange,
        string? locationRange, string type, bool? isVertical)
    {
        if (dataRange != null) parameters.Set("dataRange", dataRange);
        if (locationRange != null) parameters.Set("locationRange", locationRange);
        parameters.Set("type", type);
        if (isVertical.HasValue) parameters.Set("isVertical", isVertical.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete operation.
    /// </summary>
    private static OperationParameters BuildDeleteParameters(OperationParameters parameters, int? groupIndex)
    {
        if (groupIndex.HasValue) parameters.Set("groupIndex", groupIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the set_style operation.
    /// </summary>
    private static OperationParameters BuildSetStyleParameters(OperationParameters parameters, int? groupIndex,
        string? presetStyle, bool? showHighPoint, bool? showLowPoint, bool? showFirstPoint, bool? showLastPoint,
        bool? showNegativePoints, bool? showMarkers)
    {
        if (groupIndex.HasValue) parameters.Set("groupIndex", groupIndex.Value);
        if (presetStyle != null) parameters.Set("presetStyle", presetStyle);
        if (showHighPoint.HasValue) parameters.Set("showHighPoint", showHighPoint.Value);
        if (showLowPoint.HasValue) parameters.Set("showLowPoint", showLowPoint.Value);
        if (showFirstPoint.HasValue) parameters.Set("showFirstPoint", showFirstPoint.Value);
        if (showLastPoint.HasValue) parameters.Set("showLastPoint", showLastPoint.Value);
        if (showNegativePoints.HasValue) parameters.Set("showNegativePoints", showNegativePoints.Value);
        if (showMarkers.HasValue) parameters.Set("showMarkers", showMarkers.Value);
        return parameters;
    }
}
