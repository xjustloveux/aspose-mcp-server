using System.ComponentModel;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for Excel data operations (sort, find/replace, batch write, get content, get statistics, get used
///     range)
/// </summary>
[McpServerToolType]
public class ExcelDataOperationsTool
{
    /// <summary>
    ///     Handler registry for data operations.
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
    ///     Initializes a new instance of the <see cref="ExcelDataOperationsTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelDataOperationsTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry =
            HandlerRegistry<Workbook>.CreateFromNamespace("AsposeMcpServer.Handlers.Excel.DataOperations");
    }

    /// <summary>
    ///     Executes an Excel data operation (sort, find_replace, batch_write, get_content, get_statistics, or get_used_range).
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: sort, find_replace, batch_write, get_content, get_statistics, or
    ///     get_used_range.
    /// </param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based).</param>
    /// <param name="range">Cell range (e.g., 'A1:C10', required for sort, optional for get_content).</param>
    /// <param name="sortColumn">Column index to sort by (0-based, relative to range start, required for sort).</param>
    /// <param name="ascending">True for ascending, false for descending.</param>
    /// <param name="hasHeader">Whether the range has a header row.</param>
    /// <param name="findText">Text to find (required for find_replace).</param>
    /// <param name="replaceText">Text to replace with (required for find_replace).</param>
    /// <param name="matchCase">Match case.</param>
    /// <param name="matchEntireCell">Match entire cell content.</param>
    /// <param name="data">Data for batch_write as JSON array or object.</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "excel_data_operations")]
    [Description(
        @"Excel data operations. Supports 6 operations: sort, find_replace, batch_write, get_content, get_statistics, get_used_range.

Usage examples:
- Sort data: excel_data_operations(operation='sort', path='book.xlsx', range='A1:C10', sortColumn=0)
- Find and replace: excel_data_operations(operation='find_replace', path='book.xlsx', findText='old', replaceText='new')
- Batch write: excel_data_operations(operation='batch_write', path='book.xlsx', data=[{cell:'A1',value:'Value1'},{cell:'B1',value:'Value2'}])
- Get content: excel_data_operations(operation='get_content', path='book.xlsx', range='A1:C10')
- Get statistics: excel_data_operations(operation='get_statistics', path='book.xlsx', range='A1:A10')
- Get used range: excel_data_operations(operation='get_used_range', path='book.xlsx')")]
    public string Execute(
        [Description("Operation: sort, find_replace, batch_write, get_content, get_statistics, get_used_range")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Cell range (e.g., 'A1:C10', required for sort, optional for get_content)")]
        string? range = null,
        [Description("Column index to sort by (0-based, relative to range start, required for sort)")]
        int sortColumn = 0,
        [Description("True for ascending, false for descending (default: true)")]
        bool ascending = true,
        [Description("Whether the range has a header row (default: false)")]
        bool hasHeader = false,
        [Description("Text to find (required for find_replace)")]
        string? findText = null,
        [Description("Text to replace with (required for find_replace)")]
        string? replaceText = null,
        [Description("Match case (default: false)")]
        bool matchCase = false,
        [Description("Match entire cell content (default: false)")]
        bool matchEntireCell = false,
        [Description("Data for batch_write: [{cell:'A1',value:'val1'},...] or JSON object {A1:'val1',B1:'val2'}")]
        JsonNode? data = null)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var parameters = BuildParameters(operation, sheetIndex, range, sortColumn, ascending, hasHeader,
            findText, replaceText, matchCase, matchEntireCell, data);

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

        var op = operation.ToLowerInvariant();
        if (op == "get_content" || op == "get_statistics" || op == "get_used_range")
            return result;

        if (operationContext.IsModified)
            ctx.Save(outputPath);

        return $"{result}\n{ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Builds OperationParameters from method parameters using strategy pattern.
    ///     Parameters are documented on the Execute method.
    /// </summary>
    /// <returns>OperationParameters configured with all input values.</returns>
    private static OperationParameters BuildParameters(
        string operation,
        int sheetIndex,
        string? range,
        int sortColumn,
        bool ascending,
        bool hasHeader,
        string? findText,
        string? replaceText,
        bool matchCase,
        bool matchEntireCell,
        JsonNode? data)
    {
        var parameters = new OperationParameters();
        parameters.Set("sheetIndex", sheetIndex);

        return operation.ToLowerInvariant() switch
        {
            "sort" => BuildSortParameters(parameters, range, sortColumn, ascending, hasHeader),
            "find_replace" => BuildFindReplaceParameters(parameters, findText, replaceText, matchCase, matchEntireCell),
            "batch_write" => BuildBatchWriteParameters(parameters, data),
            "get_content" or "get_statistics" => BuildRangeParameters(parameters, range),
            _ => parameters
        };
    }

    /// <summary>
    ///     Builds parameters for the sort data operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="range">The cell range to sort.</param>
    /// <param name="sortColumn">Column index to sort by (0-based).</param>
    /// <param name="ascending">Whether to sort ascending.</param>
    /// <param name="hasHeader">Whether the range has a header row.</param>
    /// <returns>OperationParameters configured for sorting data.</returns>
    private static OperationParameters BuildSortParameters(OperationParameters parameters, string? range,
        int sortColumn, bool ascending, bool hasHeader)
    {
        if (range != null) parameters.Set("range", range);
        parameters.Set("sortColumn", sortColumn);
        parameters.Set("ascending", ascending);
        parameters.Set("hasHeader", hasHeader);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the find and replace operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="findText">The text to find.</param>
    /// <param name="replaceText">The text to replace with.</param>
    /// <param name="matchCase">Whether to match case.</param>
    /// <param name="matchEntireCell">Whether to match entire cell content.</param>
    /// <returns>OperationParameters configured for find and replace.</returns>
    private static OperationParameters BuildFindReplaceParameters(OperationParameters parameters, string? findText,
        string? replaceText, bool matchCase, bool matchEntireCell)
    {
        if (findText != null) parameters.Set("findText", findText);
        if (replaceText != null) parameters.Set("replaceText", replaceText);
        parameters.Set("matchCase", matchCase);
        parameters.Set("matchEntireCell", matchEntireCell);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters for the batch write operation.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="data">The data to write as JSON.</param>
    /// <returns>OperationParameters configured for batch writing.</returns>
    private static OperationParameters BuildBatchWriteParameters(OperationParameters parameters, JsonNode? data)
    {
        if (data != null) parameters.Set("data", data);
        return parameters;
    }

    /// <summary>
    ///     Builds parameters containing only the cell range.
    /// </summary>
    /// <param name="parameters">Base parameters with sheet index.</param>
    /// <param name="range">The cell range.</param>
    /// <returns>OperationParameters with range set.</returns>
    private static OperationParameters BuildRangeParameters(OperationParameters parameters, string? range)
    {
        if (range != null) parameters.Set("range", range);
        return parameters;
    }
}
