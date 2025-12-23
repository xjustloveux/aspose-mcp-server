using System.Text;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel hyperlinks (add, edit, delete, get)
///     Merges: ExcelAddHyperlinkTool, ExcelEditHyperlinkTool, ExcelDeleteHyperlinkTool, ExcelGetHyperlinksTool
/// </summary>
public class ExcelHyperlinkTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples
    /// </summary>
    public string Description => @"Manage Excel hyperlinks. Supports 4 operations: add, edit, delete, get.

Usage examples:
- Add hyperlink: excel_hyperlink(operation='add', path='book.xlsx', cell='A1', url='https://example.com', displayText='Link')
- Edit hyperlink: excel_hyperlink(operation='edit', path='book.xlsx', cell='A1', url='https://newurl.com')
- Delete hyperlink: excel_hyperlink(operation='delete', path='book.xlsx', cell='A1')
- Get hyperlinks: excel_hyperlink(operation='get', path='book.xlsx')";

    /// <summary>
    ///     Gets the JSON schema defining the input parameters for the tool
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
- 'edit': Edit a hyperlink (required params: path, cell, url)
- 'delete': Delete a hyperlink (required params: path, cell)
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
                    "Cell reference (e.g., 'A1', required for add, optional for edit/delete as alternative to hyperlinkIndex)"
            },
            url = new
            {
                type = "string",
                description = "URL or file path for the hyperlink (required for add)"
            },
            displayText = new
            {
                type = "string",
                description = "Display text for the hyperlink (optional)"
            },
            hyperlinkIndex = new
            {
                type = "number",
                description = "Hyperlink index (0-based, required for edit/delete, or use cell as alternative)"
            },
            address = new
            {
                type = "string",
                description = "New hyperlink address (optional, for edit)"
            },
            textToDisplay = new
            {
                type = "string",
                description = "New display text (optional, for edit)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for add/edit/delete operations, defaults to input path)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

        return operation.ToLower() switch
        {
            "add" => await AddHyperlinkAsync(arguments, path, sheetIndex),
            "edit" => await EditHyperlinkAsync(arguments, path, sheetIndex),
            "delete" => await DeleteHyperlinkAsync(arguments, path, sheetIndex),
            "get" => await GetHyperlinksAsync(arguments, path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a hyperlink to a cell
    /// </summary>
    /// <param name="arguments">JSON arguments containing cell, address, and optional screenTip, textToDisplay</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message with hyperlink details</returns>
    private Task<string> AddHyperlinkAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            var cell = ArgumentHelper.GetString(arguments, "cell");
            var url = ArgumentHelper.GetString(arguments, "url");
            var displayText = ArgumentHelper.GetStringNullable(arguments, "displayText");

            using var workbook = new Workbook(path);
            var worksheet = workbook.Worksheets[sheetIndex];
            var cellObj = worksheet.Cells[cell];

            if (!string.IsNullOrEmpty(displayText)) cellObj.PutValue(displayText);

            worksheet.Hyperlinks.Add(cell, 1, 1, url);
            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            workbook.Save(outputPath);

            return $"Hyperlink added to cell {cell}: {url}";
        });
    }

    /// <summary>
    ///     Edits an existing hyperlink
    /// </summary>
    /// <param name="arguments">JSON arguments containing cell and optional address, screenTip, textToDisplay</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message with updated hyperlink details</returns>
    private Task<string> EditHyperlinkAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            var hyperlinkIndex = ArgumentHelper.GetIntNullable(arguments, "hyperlinkIndex");
            var cell = ArgumentHelper.GetStringNullable(arguments, "cell");
            var address = ArgumentHelper.GetStringNullable(arguments, "address");
            var textToDisplay = ArgumentHelper.GetStringNullable(arguments, "textToDisplay");

            if (!hyperlinkIndex.HasValue && string.IsNullOrEmpty(cell))
                throw new ArgumentException("Either hyperlinkIndex or cell is required for edit operation");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var hyperlinks = worksheet.Hyperlinks;

            var foundIndex = hyperlinkIndex;

            // If cell is provided, find the hyperlink index by cell address
            if (!hyperlinkIndex.HasValue && !string.IsNullOrEmpty(cell))
            {
                foundIndex = null;
                int rowIndex, colIndex;
                try
                {
                    CellsHelper.CellNameToIndex(cell, out rowIndex, out colIndex);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"[ERROR] Invalid cell address format '{cell}': {ex.Message}");
                    throw new ArgumentException($"Invalid cell address: {cell}");
                }

                for (var i = 0; i < hyperlinks.Count; i++)
                {
                    var link = hyperlinks[i];
                    var area = link.Area;
                    if (rowIndex >= area.StartRow && rowIndex <= area.EndRow &&
                        colIndex >= area.StartColumn && colIndex <= area.EndColumn)
                    {
                        foundIndex = i;
                        break;
                    }
                }

                if (!foundIndex.HasValue) throw new ArgumentException($"No hyperlink found at cell {cell}");
            }

            if (!foundIndex.HasValue) throw new ArgumentException("hyperlinkIndex is required");

            var index = foundIndex.Value;
            if (index < 0 || index >= hyperlinks.Count)
                throw new ArgumentException(
                    $"Hyperlink index {index} is out of range (worksheet has {hyperlinks.Count} hyperlinks)");

            var hyperlink = hyperlinks[index];
            var oldAddress = hyperlink.Address ?? "";
            var oldText = hyperlink.TextToDisplay ?? "";

            if (!string.IsNullOrEmpty(address)) hyperlink.Address = address;

            if (!string.IsNullOrEmpty(textToDisplay)) hyperlink.TextToDisplay = textToDisplay;

            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            workbook.Save(outputPath);

            var result = $"Successfully edited hyperlink #{index}";
            if (!string.IsNullOrEmpty(cell)) result += $" (cell: {cell})";
            result += "\n";
            result += $"Old address: {oldAddress}\n";
            result += $"New address: {hyperlink.Address ?? oldAddress}\n";
            result += $"Old display text: {oldText}\n";
            result += $"New display text: {hyperlink.TextToDisplay ?? oldText}\n";
            result += $"Output: {outputPath}";

            return result;
        });
    }

    /// <summary>
    ///     Deletes a hyperlink from a cell
    /// </summary>
    /// <param name="arguments">JSON arguments containing cell</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteHyperlinkAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            var hyperlinkIndex = ArgumentHelper.GetIntNullable(arguments, "hyperlinkIndex");
            var cell = ArgumentHelper.GetStringNullable(arguments, "cell");

            if (!hyperlinkIndex.HasValue && string.IsNullOrEmpty(cell))
                throw new ArgumentException("Either hyperlinkIndex or cell is required for delete operation");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var hyperlinks = worksheet.Hyperlinks;

            var foundIndex = hyperlinkIndex;

            // If cell is provided, find the hyperlink index by cell address
            if (!hyperlinkIndex.HasValue && !string.IsNullOrEmpty(cell))
            {
                foundIndex = null;
                int rowIndex, colIndex;
                try
                {
                    CellsHelper.CellNameToIndex(cell, out rowIndex, out colIndex);
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"[ERROR] Invalid cell address format '{cell}': {ex.Message}");
                    throw new ArgumentException($"Invalid cell address: {cell}");
                }

                for (var i = 0; i < hyperlinks.Count; i++)
                {
                    var link = hyperlinks[i];
                    var area = link.Area;
                    if (rowIndex >= area.StartRow && rowIndex <= area.EndRow &&
                        colIndex >= area.StartColumn && colIndex <= area.EndColumn)
                    {
                        foundIndex = i;
                        break;
                    }
                }

                if (!foundIndex.HasValue) throw new ArgumentException($"No hyperlink found at cell {cell}");
            }

            if (!foundIndex.HasValue) throw new ArgumentException("hyperlinkIndex is required");

            var index = foundIndex.Value;
            if (index < 0 || index >= hyperlinks.Count)
                throw new ArgumentException(
                    $"Hyperlink index {index} is out of range (worksheet has {hyperlinks.Count} hyperlinks)");

            var hyperlink = hyperlinks[index];
            var address = hyperlink.Address ?? "";

            hyperlinks.RemoveAt(index);
            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            workbook.Save(outputPath);

            var remainingCount = hyperlinks.Count;

            var result = $"Successfully deleted hyperlink #{index}";
            if (!string.IsNullOrEmpty(cell)) result += $" (cell: {cell})";
            result +=
                $"\nAddress: {address}\nRemaining hyperlinks in worksheet: {remainingCount}\nOutput: {outputPath}";

            return result;
        });
    }

    /// <summary>
    ///     Gets all hyperlinks from the worksheet
    /// </summary>
    /// <param name="_">Unused parameter</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Formatted string with all hyperlinks</returns>
    private Task<string> GetHyperlinksAsync(JsonObject? _, string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var hyperlinks = worksheet.Hyperlinks;
            var result = new StringBuilder();

            result.AppendLine($"=== Hyperlink information for worksheet '{worksheet.Name}' ===\n");
            result.AppendLine($"Total hyperlinks: {hyperlinks.Count}\n");

            if (hyperlinks.Count == 0)
            {
                result.AppendLine("No hyperlinks found");
                return result.ToString();
            }

            for (var i = 0; i < hyperlinks.Count; i++)
            {
                var hyperlink = hyperlinks[i];
                result.AppendLine($"[Hyperlink {i}]");
                result.AppendLine($"Address: {hyperlink.Address ?? "(none)"}");
                result.AppendLine($"Display text: {hyperlink.TextToDisplay ?? "(none)"}");
                var area = hyperlink.Area;
                result.AppendLine(
                    $"Location: rows {area.StartRow}-{area.EndRow}, columns {area.StartColumn}-{area.EndColumn}");
                result.AppendLine();
            }

            return result.ToString();
        });
    }
}