using System.ComponentModel;
using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;
using Range = Aspose.Cells.Range;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel named ranges (add, delete, get).
/// </summary>
[McpServerToolType]
public class ExcelNamedRangeTool
{
    /// <summary>
    ///     Session identity accessor for session isolation support.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Document session manager for in-memory editing support.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelNamedRangeTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelNamedRangeTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes an Excel named range operation (add, delete, get).
    /// </summary>
    /// <param name="operation">The operation to perform: add, delete, get.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based, default: 0). Used when range does not include sheet reference.</param>
    /// <param name="name">Name for the range. Must be a valid Excel name (required for add/delete).</param>
    /// <param name="range">Cell range (e.g., 'A1:C10' or 'Sheet1!A1:C10', required for add).</param>
    /// <param name="comment">Comment for the named range (optional for add).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "excel_named_range")]
    [Description(@"Manage Excel named ranges. Supports 3 operations: add, delete, get.

Usage examples:
- Add named range: excel_named_range(operation='add', path='book.xlsx', name='MyRange', range='A1:C10')
- Add with sheet reference: excel_named_range(operation='add', path='book.xlsx', name='MyRange', range='Sheet1!A1:C10')
- Delete named range: excel_named_range(operation='delete', path='book.xlsx', name='MyRange')
- Get named ranges: excel_named_range(operation='get', path='book.xlsx')")]
    public string Execute(
        [Description("Operation: add, delete, get")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0). Used when range does not include sheet reference")]
        int sheetIndex = 0,
        [Description("Name for the range. Must be a valid Excel name (required for add/delete)")]
        string? name = null,
        [Description("Cell range (e.g., 'A1:C10' or 'Sheet1!A1:C10', required for add)")]
        string? range = null,
        [Description("Comment for the named range (optional for add)")]
        string? comment = null)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLower() switch
        {
            "add" => AddNamedRange(ctx, outputPath, sheetIndex, name, range, comment),
            "delete" => DeleteNamedRange(ctx, outputPath, name),
            "get" => GetNamedRanges(ctx),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a named range to the workbook.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The sheet index (0-based) when range does not include sheet reference.</param>
    /// <param name="name">The name for the range.</param>
    /// <param name="range">The cell range (e.g., 'A1:C10' or 'Sheet1!A1:C10').</param>
    /// <param name="comment">Optional comment for the named range.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when name or range is missing, or named range already exists.</exception>
    /// <exception cref="InvalidOperationException">Thrown when failed to create the named range.</exception>
    private static string AddNamedRange(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex, string? name,
        string? range, string? comment)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("name is required for add operation");
        if (string.IsNullOrEmpty(range))
            throw new ArgumentException("range is required for add operation");

        var workbook = ctx.Document;
        var names = workbook.Worksheets.Names;

        if (names[name] != null)
            throw new ArgumentException($"Named range '{name}' already exists.");

        try
        {
            Range rangeObject;

            if (range.Contains('!'))
            {
                rangeObject = ParseRangeWithSheetReference(workbook, range);
            }
            else
            {
                var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
                rangeObject = CreateRangeFromAddress(worksheet.Cells, range);
            }

            rangeObject.Name = name;

            var namedRange = names[name];
            if (!string.IsNullOrEmpty(comment))
                namedRange.Comment = comment;

            ctx.Save(outputPath);
            return $"Named range '{name}' added (reference: {namedRange.RefersTo}). {ctx.GetOutputMessage(outputPath)}";
        }
        catch (ArgumentException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                $"Failed to create named range '{name}' with range '{range}': {ex.Message}", ex);
        }
    }

    /// <summary>
    ///     Deletes a named range from the workbook.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="name">The name of the range to delete.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when name is missing or named range does not exist.</exception>
    private static string DeleteNamedRange(DocumentContext<Workbook> ctx, string? outputPath, string? name)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("name is required for delete operation");

        var workbook = ctx.Document;
        var names = workbook.Worksheets.Names;

        if (names[name] == null)
            throw new ArgumentException($"Named range '{name}' does not exist.");

        names.Remove(name);

        ctx.Save(outputPath);
        return $"Named range '{name}' deleted. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Gets all named ranges from the workbook.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <returns>A JSON string containing all named ranges information.</returns>
    private static string GetNamedRanges(DocumentContext<Workbook> ctx)
    {
        var workbook = ctx.Document;
        var names = workbook.Worksheets.Names;

        if (names.Count == 0)
        {
            var emptyResult = new
            {
                count = 0,
                items = Array.Empty<object>(),
                message = "No named ranges found"
            };
            return JsonSerializer.Serialize(emptyResult, new JsonSerializerOptions { WriteIndented = true });
        }

        List<object> nameList = [];
        for (var i = 0; i < names.Count; i++)
        {
            var namedRange = names[i];
            nameList.Add(new
            {
                index = i,
                name = namedRange.Text,
                reference = namedRange.RefersTo,
                comment = namedRange.Comment,
                isVisible = namedRange.IsVisible
            });
        }

        var result = new
        {
            count = names.Count,
            items = nameList
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Parses a range address that includes a sheet reference (e.g., "Sheet1!A1:B2").
    /// </summary>
    /// <param name="workbook">The workbook containing the worksheets.</param>
    /// <param name="rangeAddress">The range address with sheet reference (e.g., 'Sheet1!A1:B2').</param>
    /// <returns>The Range object corresponding to the address.</returns>
    /// <exception cref="ArgumentException">Thrown when the range format is invalid or the worksheet is not found.</exception>
    private static Range ParseRangeWithSheetReference(Workbook workbook, string rangeAddress)
    {
        var exclamationIndex = rangeAddress.LastIndexOf('!');
        if (exclamationIndex <= 0)
            throw new ArgumentException($"Invalid range format: '{rangeAddress}'. Expected format: 'SheetName!A1:C1'");

        var sheetRef = rangeAddress[..exclamationIndex].Trim().Trim('\'');
        var cellRange = rangeAddress[(exclamationIndex + 1)..].Trim();

        Worksheet? targetSheet = null;
        foreach (var ws in workbook.Worksheets)
            if (ws.Name == sheetRef)
            {
                targetSheet = ws;
                break;
            }

        if (targetSheet == null)
            throw new ArgumentException($"Worksheet '{sheetRef}' not found.");

        return CreateRangeFromAddress(targetSheet.Cells, cellRange);
    }

    /// <summary>
    ///     Creates a Range object from a cell address (e.g., "A1:B2" or "A1").
    /// </summary>
    /// <param name="cells">The Cells collection of the worksheet.</param>
    /// <param name="address">The cell address (e.g., 'A1:B2' or 'A1').</param>
    /// <returns>The Range object corresponding to the address.</returns>
    private static Range CreateRangeFromAddress(Cells cells, string address)
    {
        var colonIndex = address.IndexOf(':');
        if (colonIndex > 0)
        {
            var startCell = address[..colonIndex].Trim();
            var endCell = address[(colonIndex + 1)..].Trim();
            return cells.CreateRange(startCell, endCell);
        }

        return cells.CreateRange(address.Trim(), address.Trim());
    }
}