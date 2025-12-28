using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel hyperlinks (add, edit, delete, get).
/// </summary>
public class ExcelHyperlinkTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples.
    /// </summary>
    public string Description => @"Manage Excel hyperlinks. Supports 4 operations: add, edit, delete, get.

Usage examples:
- Add hyperlink: excel_hyperlink(operation='add', path='book.xlsx', cell='A1', url='https://example.com', displayText='Link')
- Edit hyperlink: excel_hyperlink(operation='edit', path='book.xlsx', cell='A1', url='https://newurl.com')
- Delete hyperlink: excel_hyperlink(operation='delete', path='book.xlsx', cell='A1')
- Get hyperlinks: excel_hyperlink(operation='get', path='book.xlsx')";

    /// <summary>
    ///     Gets the JSON schema defining the input parameters for the tool.
    /// </summary>
    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add a hyperlink (required params: path, cell, url)
- 'edit': Edit a hyperlink (required params: path, cell or hyperlinkIndex, url)
- 'delete': Delete a hyperlink (required params: path, cell or hyperlinkIndex)
- 'get': Get all hyperlinks (required params: path)",
                @enum = new[] { "add", "edit", "delete", "get" }
            },
            path = new
            {
                type = "string",
                description = "Excel file path (required for all operations)"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, default: 0)"
            },
            cell = new
            {
                type = "string",
                description =
                    "Cell reference in A1 notation (e.g., 'A1', 'B2'). Required for add, optional for edit/delete."
            },
            url = new
            {
                type = "string",
                description = "URL or file path for the hyperlink (required for add, optional for edit)"
            },
            displayText = new
            {
                type = "string",
                description = "Display text for the hyperlink (optional for add/edit)"
            },
            hyperlinkIndex = new
            {
                type = "number",
                description = "Hyperlink index (0-based, alternative to cell for edit/delete)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to input path)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments.
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters.</param>
    /// <returns>Result message as a string.</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

        return operation.ToLower() switch
        {
            "add" => await AddHyperlinkAsync(path, outputPath, sheetIndex, arguments),
            "edit" => await EditHyperlinkAsync(path, outputPath, sheetIndex, arguments),
            "delete" => await DeleteHyperlinkAsync(path, outputPath, sheetIndex, arguments),
            "get" => await GetHyperlinksAsync(path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a hyperlink to a cell.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing cell, url, and optional displayText.</param>
    /// <returns>Success message with hyperlink details.</returns>
    /// <exception cref="ArgumentException">Thrown if required parameters are missing or cell already has a hyperlink.</exception>
    private Task<string> AddHyperlinkAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var cell = ArgumentHelper.GetString(arguments, "cell");
            var url = ArgumentHelper.GetString(arguments, "url");
            var displayText = ArgumentHelper.GetStringNullable(arguments, "displayText");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            var existingIndex = FindHyperlinkIndexByCell(worksheet.Hyperlinks, cell);
            if (existingIndex.HasValue)
                throw new ArgumentException($"Cell {cell} already has a hyperlink. Use 'edit' operation to modify it.");

            var hyperlinkIndex = worksheet.Hyperlinks.Add(cell, 1, 1, url);
            var hyperlink = worksheet.Hyperlinks[hyperlinkIndex];

            if (!string.IsNullOrEmpty(displayText))
                hyperlink.TextToDisplay = displayText;

            workbook.Save(outputPath);

            return $"Hyperlink added to {cell}: {url}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits an existing hyperlink.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing cell or hyperlinkIndex, and optional url, displayText.</param>
    /// <returns>Success message with updated hyperlink details.</returns>
    /// <exception cref="ArgumentException">Thrown if hyperlink not found or parameters invalid.</exception>
    private Task<string> EditHyperlinkAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var hyperlinkIndex = ArgumentHelper.GetIntNullable(arguments, "hyperlinkIndex");
            var cell = ArgumentHelper.GetStringNullable(arguments, "cell");
            var url = ArgumentHelper.GetStringNullable(arguments, "url");
            var displayText = ArgumentHelper.GetStringNullable(arguments, "displayText");

            if (!hyperlinkIndex.HasValue && string.IsNullOrEmpty(cell))
                throw new ArgumentException("Either 'hyperlinkIndex' or 'cell' is required for edit operation.");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var hyperlinks = worksheet.Hyperlinks;

            var index = ResolveHyperlinkIndex(hyperlinks, hyperlinkIndex, cell);
            var hyperlink = hyperlinks[index];
            var changes = new List<string>();

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

            workbook.Save(outputPath);

            var cellRef = CellsHelper.CellIndexToName(hyperlink.Area.StartRow, hyperlink.Area.StartColumn);
            return changes.Count > 0
                ? $"Hyperlink at {cellRef} edited: {string.Join(", ", changes)}. Output: {outputPath}"
                : $"Hyperlink at {cellRef} unchanged. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes a hyperlink from a cell.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <param name="arguments">JSON arguments containing cell or hyperlinkIndex.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown if hyperlink not found or parameters invalid.</exception>
    private Task<string> DeleteHyperlinkAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var hyperlinkIndex = ArgumentHelper.GetIntNullable(arguments, "hyperlinkIndex");
            var cell = ArgumentHelper.GetStringNullable(arguments, "cell");

            if (!hyperlinkIndex.HasValue && string.IsNullOrEmpty(cell))
                throw new ArgumentException("Either 'hyperlinkIndex' or 'cell' is required for delete operation.");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var hyperlinks = worksheet.Hyperlinks;

            var index = ResolveHyperlinkIndex(hyperlinks, hyperlinkIndex, cell);
            var cellRef =
                CellsHelper.CellIndexToName(hyperlinks[index].Area.StartRow, hyperlinks[index].Area.StartColumn);

            hyperlinks.RemoveAt(index);
            workbook.Save(outputPath);

            return $"Hyperlink at {cellRef} deleted. {hyperlinks.Count} hyperlinks remaining. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all hyperlinks from the worksheet.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based).</param>
    /// <returns>JSON string with all hyperlinks.</returns>
    private Task<string> GetHyperlinksAsync(string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            using var workbook = new Workbook(path);
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

            var hyperlinkList = new List<object>();
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
        });
    }

    /// <summary>
    ///     Finds hyperlink index by cell reference.
    /// </summary>
    /// <param name="hyperlinks">Hyperlink collection.</param>
    /// <param name="cell">Cell reference in A1 notation.</param>
    /// <returns>Hyperlink index if found, null otherwise.</returns>
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
    /// <param name="hyperlinks">Hyperlink collection.</param>
    /// <param name="hyperlinkIndex">Direct hyperlink index (optional).</param>
    /// <param name="cell">Cell reference (optional).</param>
    /// <returns>Resolved hyperlink index.</returns>
    /// <exception cref="ArgumentException">Thrown if hyperlink not found or index out of range.</exception>
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