using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;
using Range = Aspose.Cells.Range;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel named ranges (add, delete, get).
/// </summary>
public class ExcelNamedRangeTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples.
    /// </summary>
    public string Description => @"Manage Excel named ranges. Supports 3 operations: add, delete, get.

Usage examples:
- Add named range: excel_named_range(operation='add', path='book.xlsx', name='MyRange', range='A1:C10')
- Add with sheet reference: excel_named_range(operation='add', path='book.xlsx', name='MyRange', range='Sheet1!A1:C10')
- Delete named range: excel_named_range(operation='delete', path='book.xlsx', name='MyRange')
- Get named ranges: excel_named_range(operation='get', path='book.xlsx')";

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
- 'add': Add a named range (required params: path, name, range)
- 'delete': Delete a named range (required params: path, name)
- 'get': Get all named ranges (required params: path)",
                @enum = new[] { "add", "delete", "get" }
            },
            path = new
            {
                type = "string",
                description = "Excel file path (required for all operations)"
            },
            name = new
            {
                type = "string",
                description = "Name for the range. Must be a valid Excel name (required for add/delete)"
            },
            range = new
            {
                type = "string",
                description = "Cell range (e.g., 'A1:C10' or 'Sheet1!A1:C10', required for add)"
            },
            comment = new
            {
                type = "string",
                description = "Comment for the named range (optional for add)"
            },
            sheetIndex = new
            {
                type = "number",
                description =
                    "Sheet index (0-based, optional, default: 0). Used when range does not include sheet reference"
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
    /// <exception cref="ArgumentException">Thrown when operation is unknown.</exception>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

        return operation.ToLower() switch
        {
            "add" => await AddNamedRangeAsync(path, outputPath, sheetIndex, arguments),
            "delete" => await DeleteNamedRangeAsync(path, outputPath, arguments),
            "get" => await GetNamedRangesAsync(path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a named range to the workbook.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="sheetIndex">Worksheet index (0-based), used when range has no sheet reference.</param>
    /// <param name="arguments">JSON arguments containing name, range, and optional comment.</param>
    /// <returns>Success message with named range details.</returns>
    /// <exception cref="ArgumentException">Thrown when named range already exists or range format is invalid.</exception>
    /// <exception cref="InvalidOperationException">Thrown when named range creation fails.</exception>
    private Task<string> AddNamedRangeAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var name = ArgumentHelper.GetString(arguments, "name");
            var rangeAddress = ArgumentHelper.GetString(arguments, "range");
            var comment = ArgumentHelper.GetStringNullable(arguments, "comment");

            using var workbook = new Workbook(path);
            var names = workbook.Worksheets.Names;

            if (names[name] != null)
                throw new ArgumentException($"Named range '{name}' already exists.");

            try
            {
                Range rangeObject;

                if (rangeAddress.Contains('!'))
                {
                    rangeObject = ParseRangeWithSheetReference(workbook, rangeAddress);
                }
                else
                {
                    var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
                    rangeObject = CreateRangeFromAddress(worksheet.Cells, rangeAddress);
                }

                rangeObject.Name = name;

                var namedRange = names[name];
                if (!string.IsNullOrEmpty(comment))
                    namedRange.Comment = comment;

                workbook.Save(outputPath);

                return $"Named range '{name}' added (reference: {namedRange.RefersTo}). Output: {outputPath}";
            }
            catch (ArgumentException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"Failed to create named range '{name}' with range '{rangeAddress}': {ex.Message}", ex);
            }
        });
    }

    /// <summary>
    ///     Deletes a named range from the workbook.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing name.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when named range does not exist.</exception>
    private Task<string> DeleteNamedRangeAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var name = ArgumentHelper.GetString(arguments, "name");

            using var workbook = new Workbook(path);
            var names = workbook.Worksheets.Names;

            if (names[name] == null)
                throw new ArgumentException($"Named range '{name}' does not exist.");

            names.Remove(name);
            workbook.Save(outputPath);

            return $"Named range '{name}' deleted. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all named ranges from the workbook.
    /// </summary>
    /// <param name="path">Excel file path.</param>
    /// <returns>JSON string containing all named ranges.</returns>
    private Task<string> GetNamedRangesAsync(string path)
    {
        return Task.Run(() =>
        {
            using var workbook = new Workbook(path);
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

            var nameList = new List<object>();
            for (var i = 0; i < names.Count; i++)
            {
                var name = names[i];
                nameList.Add(new
                {
                    index = i,
                    name = name.Text,
                    reference = name.RefersTo,
                    comment = name.Comment,
                    isVisible = name.IsVisible
                });
            }

            var result = new
            {
                count = names.Count,
                items = nameList
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }

    /// <summary>
    ///     Parses a range address that includes a sheet reference (e.g., "Sheet1!A1:B2").
    /// </summary>
    /// <param name="workbook">The workbook containing the worksheet.</param>
    /// <param name="rangeAddress">Range address with sheet reference.</param>
    /// <returns>The created Range object.</returns>
    /// <exception cref="ArgumentException">Thrown when format is invalid or worksheet not found.</exception>
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
    /// <param name="cells">The Cells collection to create the range from.</param>
    /// <param name="address">Cell address in A1:B2 or A1 format.</param>
    /// <returns>The created Range object.</returns>
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