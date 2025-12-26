using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;
using Range = Aspose.Cells.Range;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel named ranges (add, delete, get)
///     Merges: ExcelAddNamedRangeTool, ExcelDeleteNamedRangeTool, ExcelGetNamedRangesTool
/// </summary>
public class ExcelNamedRangeTool : IAsposeTool
{
    public string Description => @"Manage Excel named ranges. Supports 3 operations: add, delete, get.

Usage examples:
- Add named range: excel_named_range(operation='add', path='book.xlsx', name='MyRange', range='A1:C10')
- Delete named range: excel_named_range(operation='delete', path='book.xlsx', name='MyRange')
- Get named ranges: excel_named_range(operation='get', path='book.xlsx')";

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
                description = "Name for the range (required for add/delete)"
            },
            range = new
            {
                type = "string",
                description = "Cell range (e.g., 'A1:C10') or formula (required for add)"
            },
            comment = new
            {
                type = "string",
                description = "Comment for the named range (optional, for add)"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, default: 0, for add operation)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for add/delete operations, defaults to input path)"
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
    ///     Adds a named range to the workbook
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing name and range</param>
    /// <returns>Success message with named range details</returns>
    private Task<string> AddNamedRangeAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var name = ArgumentHelper.GetString(arguments, "name");
            var range = ArgumentHelper.GetString(arguments, "range");
            var comment = ArgumentHelper.GetStringNullable(arguments, "comment");

            using var workbook = new Workbook(path);
            var names = workbook.Worksheets.Names;

            // Get the correct worksheet using sheetIndex
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            // Use Range object approach instead of manually constructing RefersTo string
            // This is the recommended way according to Aspose.Cells documentation
            // It automatically handles sheet name escaping, absolute references, and special characters
            try
            {
                // Check if name already exists using get operation's method
                using var checkWorkbook = new Workbook(path);
                var checkNames = checkWorkbook.Worksheets.Names;
                foreach (var checkName in checkNames)
                    try
                    {
                        var checkText = checkName.Text;
                        if (!string.IsNullOrEmpty(checkText) && checkText == name)
                            throw new ArgumentException($"Named range '{name}' already exists");
                    }
                    catch (ArgumentException)
                    {
                        throw;
                    }
                    catch (Exception ex)
                    {
                        // Ignore if named range creation fails (will be handled by error handling)
                        Console.Error.WriteLine(
                            $"[WARN] Named range creation failed (will be handled by error handling): {ex.Message}");
                    }

                // Use Range object approach instead of manually constructing RefersTo string
                // This is the recommended way according to Aspose.Cells documentation
                // It automatically handles sheet name escaping, absolute references, and special characters

                Range rangeObject;
                if (range.Contains("!"))
                {
                    // Range already contains sheet reference, parse it
                    var parts = range.Split('!');
                    if (parts.Length == 2)
                    {
                        var sheetRef = parts[0].Trim().Trim('\'');
                        var cellRange = parts[1].Trim();

                        // Find the worksheet by name
                        Worksheet? targetSheet = null;
                        foreach (var ws in workbook.Worksheets)
                            if (ws.Name == sheetRef)
                            {
                                targetSheet = ws;
                                break;
                            }

                        if (targetSheet == null)
                            throw new ArgumentException($"Worksheet '{sheetRef}' not found");

                        // Parse cell range (e.g., "A1:C1")
                        var cellParts = cellRange.Split(':');
                        rangeObject = targetSheet.Cells.CreateRange(cellParts[0].Trim(),
                            cellParts.Length == 2 ? cellParts[1].Trim() : cellParts[0].Trim());
                    }
                    else
                    {
                        throw new ArgumentException(
                            $"Invalid range format with sheet reference: '{range}'. Expected format: 'SheetName!A1:C1' or 'SheetName!A1'");
                    }
                }
                else
                {
                    // Range without sheet reference, use the specified worksheet
                    // Parse cell range (e.g., "A1:C1")
                    var cellParts = range.Split(':');
                    rangeObject = worksheet.Cells.CreateRange(cellParts[0].Trim(),
                        cellParts.Length == 2 ? cellParts[1].Trim() : range.Trim());
                }

                // Set the name on the Range object - this automatically creates the named range
                // with correct RefersTo format including sheet name, absolute references, etc.
                rangeObject.Name = name;

                // Get the actual named range to verify and get RefersTo
                var namedRange = names[name];
                var actualRefersTo = namedRange.RefersTo;

                // Set comment if provided
                if (!string.IsNullOrEmpty(comment)) namedRange.Comment = comment;

                // Save the workbook to persist the changes
                workbook.Save(outputPath);

                return $"Named range '{name}' added with reference {actualRefersTo}. Output: {outputPath}";
            }
            catch (ArgumentException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"Unable to create named range '{name}', range: {range}. Error: {ex.Message}", ex);
            }
        });
    }

    /// <summary>
    ///     Deletes a named range from the workbook
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing name</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteNamedRangeAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var name = ArgumentHelper.GetString(arguments, "name");

            using var workbook = new Workbook(path);
            var names = workbook.Worksheets.Names;

            Name? namedRange;
            try
            {
                namedRange = names[name];
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[ERROR] Error accessing named range '{name}': {ex.Message}");
                throw new ArgumentException($"Named range '{name}' does not exist");
            }

            if (namedRange == null) throw new ArgumentException($"Named range '{name}' does not exist");

            // Find the index of the named range
            var indexToRemove = -1;
            for (var i = 0; i < names.Count; i++)
                if (names[i] == namedRange)
                {
                    indexToRemove = i;
                    break;
                }

            if (indexToRemove >= 0) names.RemoveAt(indexToRemove);
            workbook.Save(outputPath);

            return $"Named range '{name}' deleted. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all named ranges from the workbook
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <returns>JSON string with all named ranges</returns>
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
}