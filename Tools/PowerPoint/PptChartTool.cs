using System.ComponentModel;
using System.Text.Json.Nodes;
using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint charts (add, edit, delete, get data, update data)
///     Merges: PptAddChartTool, PptEditChartTool, PptDeleteChartTool, PptGetChartDataTool, PptUpdateChartDataTool
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.PowerPoint.Chart")]
[McpServerToolType]
public class PptChartTool
{
    /// <summary>
    ///     Handler registry for chart operations.
    /// </summary>
    private readonly HandlerRegistry<Presentation> _handlerRegistry;

    /// <summary>
    ///     Identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document lifecycle management.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptChartTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptChartTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Presentation>.CreateFromNamespace("AsposeMcpServer.Handlers.PowerPoint.Chart");
    }

    /// <summary>
    ///     Executes a PowerPoint chart operation (add, edit, delete, get_data, update_data).
    /// </summary>
    /// <param name="operation">The operation to perform: add, edit, delete, get_data, update_data.</param>
    /// <param name="slideIndex">Slide index (0-based).</param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="shapeIndex">Chart index (0-based, required for edit/delete/get_data/update_data).</param>
    /// <param name="chartType">Chart type (Column, Bar, Line, Pie, etc., required for add, optional for edit).</param>
    /// <param name="title">Chart title (optional).</param>
    /// <param name="x">Chart X position in points (optional for add, default: 50).</param>
    /// <param name="y">Chart Y position in points (optional for add, default: 50).</param>
    /// <param name="width">Chart width in points (optional for add, default: 500).</param>
    /// <param name="height">Chart height in points (optional for add, default: 400).</param>
    /// <param name="data">Chart data object with categories and series (optional, for edit/update_data).</param>
    /// <param name="clearExisting">Clear existing data before adding new (optional, for update_data, default: false).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "ppt_chart",
        Title = "PowerPoint Chart Operations",
        Destructive = true,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(@"Manage PowerPoint charts. Supports 5 operations: add, edit, delete, get_data, update_data.

Usage examples:
- Add chart: ppt_chart(operation='add', path='presentation.pptx', slideIndex=0, chartType='Column', x=100, y=100, width=400, height=300)
- Edit chart: ppt_chart(operation='edit', path='presentation.pptx', slideIndex=0, shapeIndex=0, title='New Title')
- Delete chart: ppt_chart(operation='delete', path='presentation.pptx', slideIndex=0, shapeIndex=0)
- Get data: ppt_chart(operation='get_data', path='presentation.pptx', slideIndex=0, shapeIndex=0)
- Update data: ppt_chart(operation='update_data', path='presentation.pptx', slideIndex=0, shapeIndex=0, data={categories:['A','B'],series:[{name:'Sales',values:[1,2]}]})

Note: shapeIndex refers to the chart index (0-based) among all charts on the slide, not the absolute shape index.")]
    public object Execute(
        [Description("Operation: add, edit, delete, get_data, update_data")]
        string operation,
        [Description("Slide index (0-based)")] int slideIndex,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Chart index (0-based, required for edit/delete/get_data/update_data)")]
        int? shapeIndex = null,
        [Description("Chart type (Column, Bar, Line, Pie, etc., required for add, optional for edit)")]
        string? chartType = null,
        [Description("Chart title (optional)")]
        string? title = null,
        [Description("Chart X position in points (optional for add, default: 50)")]
        float x = 50,
        [Description("Chart Y position in points (optional for add, default: 50)")]
        float y = 50,
        [Description("Chart width in points (optional for add, default: 500)")]
        float width = 500,
        [Description("Chart height in points (optional for add, default: 400)")]
        float height = 400,
        [Description("Chart data object with categories and series (optional, for edit/update_data)")]
        JsonObject? data = null,
        [Description("Clear existing data before adding new (optional, for update_data, default: false)")]
        bool clearExisting = false)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, slideIndex, shapeIndex, chartType, title,
            x, y, width, height, data, clearExisting);

        var handler = _handlerRegistry.GetHandler(operation);

        var operationContext = new OperationContext<Presentation>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = outputPath
        };

        var result = handler.Execute(operationContext, parameters);

        if (string.Equals(operation, "get_data", StringComparison.OrdinalIgnoreCase))
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
        int slideIndex,
        int? shapeIndex,
        string? chartType,
        string? title,
        float x,
        float y,
        float width,
        float height,
        JsonObject? data,
        bool clearExisting)
    {
        var parameters = new OperationParameters();
        parameters.Set("slideIndex", slideIndex);

        return operation.ToLowerInvariant() switch
        {
            "add" => BuildAddParameters(parameters, chartType, title, x, y, width, height),
            "edit" => BuildEditParameters(parameters, shapeIndex, title, chartType),
            "delete" => BuildDeleteParameters(parameters, shapeIndex),
            "get_data" => BuildGetDataParameters(parameters, shapeIndex),
            "update_data" => BuildUpdateDataParameters(parameters, shapeIndex, data, clearExisting),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the add chart operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters with slide index.</param>
    /// <param name="chartType">The chart type (Column, Bar, Line, Pie, etc.).</param>
    /// <param name="title">The chart title.</param>
    /// <param name="x">The X position in points.</param>
    /// <param name="y">The Y position in points.</param>
    /// <param name="width">The width in points.</param>
    /// <param name="height">The height in points.</param>
    /// <returns>OperationParameters configured for adding a chart.</returns>
    private static OperationParameters BuildAddParameters(OperationParameters parameters, string? chartType,
        string? title, float x, float y, float width, float height)
    {
        if (chartType != null) parameters.Set("chartType", chartType);
        if (title != null) parameters.Set("title", title);
        parameters.Set("x", x);
        parameters.Set("y", y);
        parameters.Set("width", width);
        parameters.Set("height", height);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the edit chart operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters with slide index.</param>
    /// <param name="shapeIndex">The chart index (0-based).</param>
    /// <param name="title">The new chart title.</param>
    /// <param name="chartType">The new chart type.</param>
    /// <returns>OperationParameters configured for editing a chart.</returns>
    private static OperationParameters BuildEditParameters(OperationParameters parameters, int? shapeIndex,
        string? title, string? chartType)
    {
        if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
        if (title != null) parameters.Set("title", title);
        if (chartType != null) parameters.Set("chartType", chartType);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the delete chart operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters with slide index.</param>
    /// <param name="shapeIndex">The chart index (0-based).</param>
    /// <returns>OperationParameters configured for deleting a chart.</returns>
    private static OperationParameters BuildDeleteParameters(OperationParameters parameters, int? shapeIndex)
    {
        if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the get chart data operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters with slide index.</param>
    /// <param name="shapeIndex">The chart index (0-based).</param>
    /// <returns>OperationParameters configured for getting chart data.</returns>
    private static OperationParameters BuildGetDataParameters(OperationParameters parameters, int? shapeIndex)
    {
        return BuildDeleteParameters(parameters, shapeIndex);
    }

    /// <summary>
    ///     Builds parameters for the update chart data operation.
    /// </summary>
    /// <param name="parameters">The base operation parameters with slide index.</param>
    /// <param name="shapeIndex">The chart index (0-based).</param>
    /// <param name="data">The chart data object with categories and series.</param>
    /// <param name="clearExisting">Whether to clear existing data before adding new.</param>
    /// <returns>OperationParameters configured for updating chart data.</returns>
    private static OperationParameters BuildUpdateDataParameters(OperationParameters parameters, int? shapeIndex,
        JsonObject? data, bool clearExisting)
    {
        if (shapeIndex.HasValue) parameters.Set("shapeIndex", shapeIndex.Value);
        if (data != null) parameters.Set("data", data);
        parameters.Set("clearExisting", clearExisting);
        return parameters;
    }
}
