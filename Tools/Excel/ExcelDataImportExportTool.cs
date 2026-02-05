using System.ComponentModel;
using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Helpers;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for Excel data import/export operations (import_json, import_array, export_csv, export_range_image).
/// </summary>
[ToolHandlerMapping("AsposeMcpServer.Handlers.Excel.DataImportExport")]
[McpServerToolType]
public class ExcelDataImportExportTool
{
    /// <summary>
    ///     Handler registry for data import/export operations.
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
    ///     Initializes a new instance of the <see cref="ExcelDataImportExportTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelDataImportExportTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.DataImportExport");
    }

    /// <summary>
    ///     Executes an Excel data import/export operation.
    /// </summary>
    /// <param name="operation">The operation to perform.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (for export/save operations).</param>
    /// <param name="sheetIndex">Sheet index (0-based).</param>
    /// <param name="jsonData">JSON data string (for import_json).</param>
    /// <param name="arrayData">Array data as semicolon-separated rows, comma-separated values (for import_array).</param>
    /// <param name="startCell">Starting cell for import (default: 'A1').</param>
    /// <param name="isVertical">Import array vertically (for import_array, default: false).</param>
    /// <param name="separator">CSV separator character (for export_csv, default: ',').</param>
    /// <param name="format">Image format for export (for export_range_image, default: 'png').</param>
    /// <param name="dpi">DPI for image export (for export_range_image, default: 150).</param>
    /// <returns>A message or data indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(
        Name = "excel_data_import_export",
        Title = "Excel Data Import/Export Operations",
        Destructive = true,
        Idempotent = false,
        OpenWorld = false,
        ReadOnly = false,
        UseStructuredContent = true)]
    [Description(
        @"Import and export Excel data. Supports 4 operations: import_json, import_array, export_csv, export_range_image.

Usage examples:
- Import JSON: excel_data_import_export(operation='import_json', path='book.xlsx', jsonData='[{""name"":""John"",""age"":30}]', outputPath='out.xlsx')
- Import array: excel_data_import_export(operation='import_array', path='book.xlsx', arrayData='A,B,C;1,2,3;4,5,6', outputPath='out.xlsx')
- Export CSV: excel_data_import_export(operation='export_csv', path='book.xlsx', outputPath='data.csv')
- Export image: excel_data_import_export(operation='export_range_image', path='book.xlsx', outputPath='sheet.png')")]
    public object Execute(
        [Description(@"Operation to perform.
- 'import_json': Import JSON data (required params: jsonData)
- 'import_array': Import array data (required params: arrayData)
- 'export_csv': Export worksheet to CSV (required params: outputPath)
- 'export_range_image': Export worksheet to image (required params: outputPath)")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path")] string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("JSON data string (for import_json)")]
        string? jsonData = null,
        [Description("Array data: semicolon-separated rows, comma-separated values (for import_array)")]
        string? arrayData = null,
        [Description("Starting cell for import (default: 'A1')")]
        string startCell = "A1",
        [Description("Import array vertically (for import_array, default: false)")]
        bool isVertical = false,
        [Description("CSV separator character (for export_csv, default: ',')")]
        string separator = ",",
        [Description("Image format (for export_range_image: png, jpeg, bmp, tiff, svg; default: png)")]
        string format = "png",
        [Description("DPI for image export (for export_range_image, default: 150)")]
        int dpi = 150)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, sheetIndex, jsonData, arrayData, startCell, isVertical,
            outputPath, separator, format, dpi);

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
        string operation, int sheetIndex, string? jsonData, string? arrayData, string startCell,
        bool isVertical, string? outputPath, string separator, string format, int dpi)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);

        return operation.ToLowerInvariant() switch
        {
            "import_json" => BuildImportJsonParameters(parameters, jsonData, startCell),
            "import_array" => BuildImportArrayParameters(parameters, arrayData, startCell, isVertical),
            "export_csv" => BuildExportCsvParameters(parameters, outputPath, separator),
            "export_range_image" => BuildExportRangeImageParameters(parameters, outputPath, format, dpi),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the import_json operation.
    /// </summary>
    private static OperationParameters BuildImportJsonParameters(OperationParameters parameters, string? jsonData,
        string startCell)
    {
        if (jsonData != null) parameters.Set("jsonData", jsonData);
        parameters.Set("startCell", startCell);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the import_array operation.
    /// </summary>
    private static OperationParameters BuildImportArrayParameters(OperationParameters parameters, string? arrayData,
        string startCell, bool isVertical)
    {
        if (arrayData != null) parameters.Set("arrayData", arrayData);
        parameters.Set("startCell", startCell);
        parameters.Set("isVertical", isVertical);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the export_csv operation.
    /// </summary>
    private static OperationParameters BuildExportCsvParameters(OperationParameters parameters, string? outputPath,
        string separator)
    {
        if (outputPath != null) parameters.Set("outputPath", outputPath);
        parameters.Set("separator", separator);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the export_range_image operation.
    /// </summary>
    private static OperationParameters BuildExportRangeImageParameters(OperationParameters parameters,
        string? outputPath, string format, int dpi)
    {
        if (outputPath != null) parameters.Set("outputPath", outputPath);
        parameters.Set("format", format);
        parameters.Set("dpi", dpi);
        return parameters;
    }
}
