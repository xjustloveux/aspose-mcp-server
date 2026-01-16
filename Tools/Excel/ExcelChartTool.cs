using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel charts (add, edit, delete, get, update data, set properties)
/// </summary>
[McpServerToolType]
public class ExcelChartTool
{
    /// <summary>
    ///     Handler registry for chart operations.
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
    ///     Initializes a new instance of the <see cref="ExcelChartTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelChartTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.Chart");
    }

    /// <summary>
    ///     Executes an Excel chart operation (add, edit, delete, get, update_data, or set_properties).
    /// </summary>
    /// <param name="operation">The operation to perform: add, edit, delete, get, update_data, or set_properties.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based).</param>
    /// <param name="chartIndex">Chart index (0-based, required for edit/delete/update_data/set_properties).</param>
    /// <param name="chartType">Chart type: Column, Bar, Line, Pie, Area, Scatter, etc.</param>
    /// <param name="dataRange">Data range for chart values.</param>
    /// <param name="categoryAxisDataRange">Category axis (X-axis) data range.</param>
    /// <param name="title">Chart title.</param>
    /// <param name="topRow">Top row index for chart position (0-based).</param>
    /// <param name="leftColumn">Left column index for chart position (0-based).</param>
    /// <param name="width">Chart width in columns.</param>
    /// <param name="height">Chart height in rows.</param>
    /// <param name="showLegend">Show legend (optional, for edit/set_properties).</param>
    /// <param name="legendPosition">Legend position: Bottom, Top, Left, Right.</param>
    /// <param name="removeTitle">Remove title (optional, for set_properties).</param>
    /// <param name="legendVisible">Legend visibility (optional, for set_properties).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "excel_chart")]
    [Description(@"Manage Excel charts. Supports 6 operations: add, edit, delete, get, update_data, set_properties.

Usage examples:
- Add chart: excel_chart(operation='add', path='book.xlsx', chartType='Column', dataRange='A1:B10', position='A12')
- Edit chart: excel_chart(operation='edit', path='book.xlsx', chartIndex=0, chartType='Line')
- Delete chart: excel_chart(operation='delete', path='book.xlsx', chartIndex=0)
- Get charts: excel_chart(operation='get', path='book.xlsx')
- Update data: excel_chart(operation='update_data', path='book.xlsx', chartIndex=0, dataRange='A1:C10')
- Set properties: excel_chart(operation='set_properties', path='book.xlsx', chartIndex=0, title='Chart Title')")]
    public string Execute(
        [Description("Operation: add, edit, delete, get, update_data, set_properties")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Chart index (0-based, required for edit/delete/update_data/set_properties)")]
        int chartIndex = 0,
        [Description(
            "Chart type: Column, Bar, Line, Pie, Area, Scatter, Doughnut, Radar, Bubble, Cylinder, Cone, Pyramid")]
        string? chartType = null,
        [Description("Data range for chart values (e.g., 'B1:B10' or 'B1:C10' for multiple series)")]
        string? dataRange = null,
        [Description("Category axis (X-axis) data range (optional, e.g., 'A1:A10')")]
        string? categoryAxisDataRange = null,
        [Description("Chart title (optional)")]
        string? title = null,
        [Description("Top row index for chart position (0-based, optional, default: auto-detect)")]
        int? topRow = null,
        [Description("Left column index for chart position (0-based, default: 0)")]
        int leftColumn = 0,
        [Description("Chart width in columns (default: 10)")]
        int width = 10,
        [Description("Chart height in rows (default: 15)")]
        int height = 15,
        [Description("Show legend (optional, for edit/set_properties)")]
        bool? showLegend = null,
        [Description("Legend position: Bottom, Top, Left, Right (optional)")]
        string? legendPosition = null,
        [Description("Remove title (optional, for set_properties)")]
        bool removeTitle = false,
        [Description("Legend visibility (optional, for set_properties)")]
        bool? legendVisible = null)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, sheetIndex, chartIndex, chartType, dataRange, categoryAxisDataRange,
            title, topRow, leftColumn, width, height, showLegend, legendPosition, removeTitle, legendVisible);

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
            return result;

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return $"{result}\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    /// </summary>
    private static OperationParameters BuildParameters(
        string operation,
        int sheetIndex,
        int chartIndex,
        string? chartType,
        string? dataRange,
        string? categoryAxisDataRange,
        string? title,
        int? topRow,
        int leftColumn,
        int width,
        int height,
        bool? showLegend,
        string? legendPosition,
        bool removeTitle,
        bool? legendVisible)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);

        return operation.ToLowerInvariant() switch
        {
            "add" => BuildAddParameters(parameters, chartType, dataRange, categoryAxisDataRange, title, topRow,
                leftColumn, width, height),
            "edit" => BuildEditParameters(parameters, chartIndex, title, dataRange, categoryAxisDataRange, chartType,
                showLegend, legendPosition),
            "delete" => BuildChartIndexParameters(parameters, chartIndex),
            "get" => parameters,
            "update_data" => BuildUpdateDataParameters(parameters, chartIndex, dataRange, categoryAxisDataRange),
            "set_properties" => BuildSetPropertiesParameters(parameters, chartIndex, title, removeTitle, legendVisible,
                legendPosition),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the add chart operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="chartType">The chart type.</param>
    /// <param name="dataRange">Data range for chart values.</param>
    /// <param name="categoryAxisDataRange">Category axis data range.</param>
    /// <param name="title">Chart title.</param>
    /// <param name="topRow">Top row index for chart position.</param>
    /// <param name="leftColumn">Left column index for chart position.</param>
    /// <param name="width">Chart width in columns.</param>
    /// <param name="height">Chart height in rows.</param>
    /// <returns>OperationParameters configured for adding a chart.</returns>
    private static OperationParameters BuildAddParameters(OperationParameters parameters, string? chartType,
        string? dataRange, string? categoryAxisDataRange, string? title, int? topRow, int leftColumn, int width,
        int height)
    {
        if (chartType != null) parameters.Set("chartType", chartType);
        if (dataRange != null) parameters.Set("dataRange", dataRange);
        if (categoryAxisDataRange != null) parameters.Set("categoryAxisDataRange", categoryAxisDataRange);
        if (title != null) parameters.Set("title", title);
        if (topRow.HasValue) parameters.Set("topRow", topRow.Value);
        parameters.Set("leftColumn", leftColumn);
        parameters.Set("width", width);
        parameters.Set("height", height);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the edit chart operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="chartIndex">The chart index (0-based).</param>
    /// <param name="title">New chart title.</param>
    /// <param name="dataRange">New data range.</param>
    /// <param name="categoryAxisDataRange">New category axis data range.</param>
    /// <param name="chartType">New chart type.</param>
    /// <param name="showLegend">Whether to show legend.</param>
    /// <param name="legendPosition">Legend position.</param>
    /// <returns>OperationParameters configured for editing a chart.</returns>
    private static OperationParameters BuildEditParameters(OperationParameters parameters, int chartIndex,
        string? title, string? dataRange, string? categoryAxisDataRange, string? chartType, bool? showLegend,
        string? legendPosition)
    {
        parameters.Set("chartIndex", chartIndex);
        if (title != null) parameters.Set("title", title);
        if (dataRange != null) parameters.Set("dataRange", dataRange);
        if (categoryAxisDataRange != null) parameters.Set("categoryAxisDataRange", categoryAxisDataRange);
        if (chartType != null) parameters.Set("chartType", chartType);
        if (showLegend.HasValue) parameters.Set("showLegend", showLegend.Value);
        if (legendPosition != null) parameters.Set("legendPosition", legendPosition);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters containing only the chart index.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="chartIndex">The chart index (0-based).</param>
    /// <returns>OperationParameters with chart index set.</returns>
    private static OperationParameters BuildChartIndexParameters(OperationParameters parameters, int chartIndex)
    {
        parameters.Set("chartIndex", chartIndex);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the update chart data operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="chartIndex">The chart index (0-based).</param>
    /// <param name="dataRange">New data range for chart values.</param>
    /// <param name="categoryAxisDataRange">New category axis data range.</param>
    /// <returns>OperationParameters configured for updating chart data.</returns>
    private static OperationParameters BuildUpdateDataParameters(OperationParameters parameters, int chartIndex,
        string? dataRange, string? categoryAxisDataRange)
    {
        parameters.Set("chartIndex", chartIndex);
        if (dataRange != null) parameters.Set("dataRange", dataRange);
        if (categoryAxisDataRange != null) parameters.Set("categoryAxisDataRange", categoryAxisDataRange);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the set chart properties operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="chartIndex">The chart index (0-based).</param>
    /// <param name="title">New chart title.</param>
    /// <param name="removeTitle">Whether to remove the title.</param>
    /// <param name="legendVisible">Legend visibility.</param>
    /// <param name="legendPosition">Legend position.</param>
    /// <returns>OperationParameters configured for setting chart properties.</returns>
    private static OperationParameters BuildSetPropertiesParameters(OperationParameters parameters, int chartIndex,
        string? title, bool removeTitle, bool? legendVisible, string? legendPosition)
    {
        parameters.Set("chartIndex", chartIndex);
        if (title != null) parameters.Set("title", title);
        parameters.Set("removeTitle", removeTitle);
        if (legendVisible.HasValue) parameters.Set("legendVisible", legendVisible.Value);
        if (legendPosition != null) parameters.Set("legendPosition", legendPosition);
        return parameters;
    }
}
