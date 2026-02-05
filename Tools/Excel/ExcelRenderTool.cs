using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Tool for rendering Excel worksheets and charts to images (render_sheet, render_chart).
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Excel.Render")]
[McpServerToolType]
public class ExcelRenderTool
{
    /// <summary>
    ///     Handler registry for render operations.
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
    ///     Initializes a new instance of the <see cref="ExcelRenderTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelRenderTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.Render");
    }

    /// <summary>
    ///     Executes an Excel render operation (render_sheet, render_chart).
    /// </summary>
    /// <param name="operation">The operation to perform.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output image file path (required).</param>
    /// <param name="sheetIndex">Sheet index (0-based).</param>
    /// <param name="chartIndex">Chart index (0-based, for render_chart).</param>
    /// <param name="format">Image format: png, jpeg, bmp, tiff, svg (default: png).</param>
    /// <param name="dpi">Rendering DPI (for render_sheet, default: 150).</param>
    /// <returns>Render result with output file paths.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "excel_render",
        Title = "Excel Render Operations",
        Destructive = false,
        Idempotent = true,
        OpenWorld = false,
        ReadOnly = true,
        UseStructuredContent = true)]
    [Description(@"Render Excel worksheets and charts to images. Supports 2 operations: render_sheet, render_chart.

Usage examples:
- Render sheet: excel_render(operation='render_sheet', path='book.xlsx', outputPath='sheet.png')
- Render with JPEG: excel_render(operation='render_sheet', path='book.xlsx', outputPath='sheet.jpg', format='jpeg', dpi=300)
- Render chart: excel_render(operation='render_chart', path='book.xlsx', outputPath='chart.png', chartIndex=0)")]
    public object Execute(
        [Description(@"Operation to perform.
- 'render_sheet': Render worksheet to image (required params: outputPath)
- 'render_chart': Render chart to image (required params: outputPath, chartIndex)")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output image file path (required)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Chart index (0-based, for render_chart)")]
        int? chartIndex = null,
        [Description("Image format: png, jpeg, bmp, tiff, svg (default: png)")]
        string format = "png",
        [Description("Rendering DPI (for render_sheet, default: 150)")]
        int dpi = 150)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, sheetIndex, outputPath, chartIndex, format, dpi);

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
        return ResultHelper.FinalizeResult((dynamic)result, ctx, outputPath);
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters(
        string operation, int sheetIndex, string? outputPath, int? chartIndex, string format, int dpi)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);
        if (outputPath != null) parameters.Set("outputPath", outputPath);
        parameters.Set("format", format);

        return operation.ToLowerInvariant() switch
        {
            "render_sheet" => BuildRenderSheetParameters(parameters, dpi),
            "render_chart" => BuildRenderChartParameters(parameters, chartIndex),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the render_sheet operation.
    /// </summary>
    private static OperationParameters BuildRenderSheetParameters(OperationParameters parameters, int dpi)
    {
        parameters.Set("dpi", dpi);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the render_chart operation.
    /// </summary>
    private static OperationParameters BuildRenderChartParameters(OperationParameters parameters, int? chartIndex)
    {
        if (chartIndex.HasValue) parameters.Set("chartIndex", chartIndex.Value);
        return parameters;
    }
}
