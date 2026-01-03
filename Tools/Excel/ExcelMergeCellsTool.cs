using System.ComponentModel;
using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel merged cells (merge, unmerge, get).
/// </summary>
[McpServerToolType]
public class ExcelMergeCellsTool
{
    /// <summary>
    ///     Document session manager for in-memory editing support.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelMergeCellsTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    public ExcelMergeCellsTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "excel_merge_cells")]
    [Description(@"Manage Excel merged cells. Supports 3 operations: merge, unmerge, get.

Usage examples:
- Merge cells: excel_merge_cells(operation='merge', path='book.xlsx', range='A1:C1')
- Unmerge cells: excel_merge_cells(operation='unmerge', path='book.xlsx', range='A1:C1')
- Get merged cells: excel_merge_cells(operation='get', path='book.xlsx')

WARNING: Merging cells will only keep the value of the top-left cell. All other cell values will be lost.")]
    public string Execute(
        [Description("Operation: merge, unmerge, get")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description(
            "Cell range to merge/unmerge (e.g., 'A1:C3', must include at least 2 cells, required for merge/unmerge)")]
        string? range = null)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path);

        return operation.ToLower() switch
        {
            "merge" => MergeCells(ctx, outputPath, sheetIndex, range),
            "unmerge" => UnmergeCells(ctx, outputPath, sheetIndex, range),
            "get" => GetMergedCells(ctx, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Merges cells in a range.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="range">The cell range to merge (e.g., 'A1:C3').</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when range is null or empty, or when the range contains only a single cell.</exception>
    private static string MergeCells(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex, string? range)
    {
        if (string.IsNullOrEmpty(range))
            throw new ArgumentException("range is required for merge operation");

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cells = worksheet.Cells;

        var cellRange = ExcelHelper.CreateRange(cells, range);

        if (cellRange is { RowCount: 1, ColumnCount: 1 })
            throw new ArgumentException(
                $"Cannot merge a single cell. Range '{range}' must include at least 2 cells.");

        cellRange.Merge();

        ctx.Save(outputPath);
        return
            $"Range {range} merged ({cellRange.RowCount} rows x {cellRange.ColumnCount} columns). {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Unmerges cells in a range.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="range">The cell range to unmerge (e.g., 'A1:C3').</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when range is null or empty.</exception>
    private static string UnmergeCells(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex, string? range)
    {
        if (string.IsNullOrEmpty(range))
            throw new ArgumentException("range is required for unmerge operation");

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cells = worksheet.Cells;

        var cellRange = ExcelHelper.CreateRange(cells, range);

        cellRange.UnMerge();

        ctx.Save(outputPath);
        return $"Range {range} unmerged. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Gets all merged cell ranges from the worksheet.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <returns>A JSON string containing the merged cells information.</returns>
    private static string GetMergedCells(DocumentContext<Workbook> ctx, int sheetIndex)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var mergedCells = worksheet.Cells.MergedCells;

        if (mergedCells == null || mergedCells.Count == 0)
        {
            var emptyResult = new
            {
                count = 0,
                worksheetName = worksheet.Name,
                items = Array.Empty<object>(),
                message = "No merged cells found"
            };
            return JsonSerializer.Serialize(emptyResult, new JsonSerializerOptions { WriteIndented = true });
        }

        List<object> mergedList = [];
        for (var i = 0; i < mergedCells.Count; i++)
        {
            var mergedCellObj = mergedCells[i];

            if (mergedCellObj is CellArea cellArea)
            {
                var startCellName = CellsHelper.CellIndexToName(cellArea.StartRow, cellArea.StartColumn);
                var endCellName = CellsHelper.CellIndexToName(cellArea.EndRow, cellArea.EndColumn);
                var rangeName = $"{startCellName}:{endCellName}";

                var cell = worksheet.Cells[cellArea.StartRow, cellArea.StartColumn];
                var cellValue = cell.Value?.ToString() ?? "(empty)";

                mergedList.Add(new
                {
                    index = i,
                    range = rangeName,
                    startCell = startCellName,
                    endCell = endCellName,
                    rowCount = cellArea.EndRow - cellArea.StartRow + 1,
                    columnCount = cellArea.EndColumn - cellArea.StartColumn + 1,
                    value = cellValue
                });
            }
        }

        var result = new
        {
            count = mergedList.Count,
            worksheetName = worksheet.Name,
            items = mergedList
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }
}