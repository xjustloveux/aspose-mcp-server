using System.ComponentModel;
using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel hyperlinks (add, edit, delete, get).
/// </summary>
[McpServerToolType]
public class ExcelHyperlinkTool
{
    /// <summary>
    ///     Document session manager for in-memory editing support.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelHyperlinkTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    public ExcelHyperlinkTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "excel_hyperlink")]
    [Description(@"Manage Excel hyperlinks. Supports 4 operations: add, edit, delete, get.

Usage examples:
- Add hyperlink: excel_hyperlink(operation='add', path='book.xlsx', cell='A1', url='https://example.com', displayText='Link')
- Edit hyperlink: excel_hyperlink(operation='edit', path='book.xlsx', cell='A1', url='https://newurl.com')
- Delete hyperlink: excel_hyperlink(operation='delete', path='book.xlsx', cell='A1')
- Get hyperlinks: excel_hyperlink(operation='get', path='book.xlsx')")]
    public string Execute(
        [Description("Operation: add, edit, delete, get")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Cell reference in A1 notation (e.g., 'A1', 'B2'). Required for add, optional for edit/delete.")]
        string? cell = null,
        [Description("URL or file path for the hyperlink (required for add, optional for edit)")]
        string? url = null,
        [Description("Display text for the hyperlink (optional for add/edit)")]
        string? displayText = null,
        [Description("Hyperlink index (0-based, alternative to cell for edit/delete)")]
        int? hyperlinkIndex = null)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path);

        return operation.ToLower() switch
        {
            "add" => AddHyperlink(ctx, outputPath, sheetIndex, cell, url, displayText),
            "edit" => EditHyperlink(ctx, outputPath, sheetIndex, cell, url, displayText, hyperlinkIndex),
            "delete" => DeleteHyperlink(ctx, outputPath, sheetIndex, cell, hyperlinkIndex),
            "get" => GetHyperlinks(ctx, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a hyperlink to a cell.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="cell">The cell reference in A1 notation.</param>
    /// <param name="url">The URL or file path for the hyperlink.</param>
    /// <param name="displayText">The display text for the hyperlink.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when cell or url is not provided, or cell already has a hyperlink.</exception>
    private static string AddHyperlink(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex, string? cell,
        string? url, string? displayText)
    {
        if (string.IsNullOrEmpty(cell))
            throw new ArgumentException("cell is required for add operation");
        if (string.IsNullOrEmpty(url))
            throw new ArgumentException("url is required for add operation");

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        var existingIndex = FindHyperlinkIndexByCell(worksheet.Hyperlinks, cell);
        if (existingIndex.HasValue)
            throw new ArgumentException($"Cell {cell} already has a hyperlink. Use 'edit' operation to modify it.");

        var hyperlinkIdx = worksheet.Hyperlinks.Add(cell, 1, 1, url);
        var hyperlink = worksheet.Hyperlinks[hyperlinkIdx];

        if (!string.IsNullOrEmpty(displayText))
            hyperlink.TextToDisplay = displayText;

        ctx.Save(outputPath);
        return $"Hyperlink added to {cell}: {url}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Edits an existing hyperlink.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="cell">The cell reference in A1 notation.</param>
    /// <param name="url">The new URL or file path for the hyperlink.</param>
    /// <param name="displayText">The new display text for the hyperlink.</param>
    /// <param name="hyperlinkIndex">The hyperlink index as an alternative to cell reference.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when neither hyperlinkIndex nor cell is provided.</exception>
    private static string EditHyperlink(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex, string? cell,
        string? url, string? displayText, int? hyperlinkIndex)
    {
        if (!hyperlinkIndex.HasValue && string.IsNullOrEmpty(cell))
            throw new ArgumentException("Either 'hyperlinkIndex' or 'cell' is required for edit operation.");

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var hyperlinks = worksheet.Hyperlinks;

        var index = ResolveHyperlinkIndex(hyperlinks, hyperlinkIndex, cell);
        var hyperlink = hyperlinks[index];
        List<string> changes = [];

        if (!string.IsNullOrEmpty(url))
        {
            hyperlink.Address = url;
            changes.Add($"url={url}");
        }

        if (!string.IsNullOrEmpty(displayText))
        {
            hyperlink.TextToDisplay = displayText;
            changes.Add($"displayText={displayText}");
        }

        ctx.Save(outputPath);

        var cellRef = CellsHelper.CellIndexToName(hyperlink.Area.StartRow, hyperlink.Area.StartColumn);
        return changes.Count > 0
            ? $"Hyperlink at {cellRef} edited: {string.Join(", ", changes)}. {ctx.GetOutputMessage(outputPath)}"
            : $"Hyperlink at {cellRef} unchanged. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Deletes a hyperlink from a cell.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="cell">The cell reference in A1 notation.</param>
    /// <param name="hyperlinkIndex">The hyperlink index as an alternative to cell reference.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when neither hyperlinkIndex nor cell is provided.</exception>
    private static string DeleteHyperlink(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        string? cell, int? hyperlinkIndex)
    {
        if (!hyperlinkIndex.HasValue && string.IsNullOrEmpty(cell))
            throw new ArgumentException("Either 'hyperlinkIndex' or 'cell' is required for delete operation.");

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var hyperlinks = worksheet.Hyperlinks;

        var index = ResolveHyperlinkIndex(hyperlinks, hyperlinkIndex, cell);
        var cellRef = CellsHelper.CellIndexToName(hyperlinks[index].Area.StartRow, hyperlinks[index].Area.StartColumn);

        hyperlinks.RemoveAt(index);

        ctx.Save(outputPath);
        return
            $"Hyperlink at {cellRef} deleted. {hyperlinks.Count} hyperlinks remaining. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Gets all hyperlinks from the worksheet.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <returns>A JSON string containing the hyperlink information.</returns>
    private static string GetHyperlinks(DocumentContext<Workbook> ctx, int sheetIndex)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var hyperlinks = worksheet.Hyperlinks;

        if (hyperlinks.Count == 0)
        {
            var emptyResult = new
            {
                count = 0,
                worksheetName = worksheet.Name,
                items = Array.Empty<object>(),
                message = "No hyperlinks found"
            };
            return JsonSerializer.Serialize(emptyResult, new JsonSerializerOptions { WriteIndented = true });
        }

        List<object> hyperlinkList = [];
        for (var i = 0; i < hyperlinks.Count; i++)
        {
            var hyperlink = hyperlinks[i];
            var area = hyperlink.Area;
            var cellRef = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);

            hyperlinkList.Add(new
            {
                index = i,
                cell = cellRef,
                url = hyperlink.Address,
                displayText = hyperlink.TextToDisplay,
                area = new
                {
                    startCell = cellRef,
                    endCell = CellsHelper.CellIndexToName(area.EndRow, area.EndColumn)
                }
            });
        }

        var result = new
        {
            count = hyperlinks.Count,
            worksheetName = worksheet.Name,
            items = hyperlinkList
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Finds hyperlink index by cell reference.
    /// </summary>
    /// <param name="hyperlinks">The hyperlink collection to search.</param>
    /// <param name="cell">The cell reference in A1 notation.</param>
    /// <returns>The hyperlink index if found, otherwise null.</returns>
    private static int? FindHyperlinkIndexByCell(HyperlinkCollection hyperlinks, string cell)
    {
        CellsHelper.CellNameToIndex(cell, out var rowIndex, out var colIndex);

        for (var i = 0; i < hyperlinks.Count; i++)
        {
            var area = hyperlinks[i].Area;
            if (rowIndex >= area.StartRow && rowIndex <= area.EndRow &&
                colIndex >= area.StartColumn && colIndex <= area.EndColumn)
                return i;
        }

        return null;
    }

    /// <summary>
    ///     Resolves hyperlink index from either direct index or cell reference.
    /// </summary>
    /// <param name="hyperlinks">The hyperlink collection to search.</param>
    /// <param name="hyperlinkIndex">The direct hyperlink index, or null to use cell reference.</param>
    /// <param name="cell">The cell reference in A1 notation as an alternative to index.</param>
    /// <returns>The resolved hyperlink index.</returns>
    /// <exception cref="ArgumentException">Thrown when neither index nor cell is provided, or hyperlink is not found.</exception>
    private static int ResolveHyperlinkIndex(HyperlinkCollection hyperlinks, int? hyperlinkIndex, string? cell)
    {
        int index;

        if (hyperlinkIndex.HasValue)
        {
            index = hyperlinkIndex.Value;
        }
        else if (!string.IsNullOrEmpty(cell))
        {
            var foundIndex = FindHyperlinkIndexByCell(hyperlinks, cell);
            if (!foundIndex.HasValue)
                throw new ArgumentException($"No hyperlink found at cell {cell}.");
            index = foundIndex.Value;
        }
        else
        {
            throw new ArgumentException("Either 'hyperlinkIndex' or 'cell' must be provided.");
        }

        if (index < 0 || index >= hyperlinks.Count)
            throw new ArgumentException(
                $"Hyperlink index {index} is out of range. Worksheet has {hyperlinks.Count} hyperlinks.");

        return index;
    }
}