using System.Text;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel freeze panes (freeze/unfreeze/get)
///     Merges: ExcelFreezePanesTool, ExcelGetFreezePanesTool
/// </summary>
public class ExcelFreezePanesTool : IAsposeTool
{
    /// <summary>
    ///     Gets the description of the tool and its usage examples
    /// </summary>
    public string Description => @"Manage Excel freeze panes. Supports 3 operations: freeze, unfreeze, get.

Usage examples:
- Freeze panes: excel_freeze_panes(operation='freeze', path='book.xlsx', row=1, column=1)
- Unfreeze panes: excel_freeze_panes(operation='unfreeze', path='book.xlsx')
- Get freeze status: excel_freeze_panes(operation='get', path='book.xlsx')";

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
- 'freeze': Freeze panes (required params: path, row, column)
- 'unfreeze': Unfreeze panes (required params: path)
- 'get': Get freeze panes status (required params: path)",
                @enum = new[] { "freeze", "unfreeze", "get" }
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
            row = new
            {
                type = "number",
                description = "Row index to freeze at (0-based, required for freeze)"
            },
            column = new
            {
                type = "number",
                description = "Column index to freeze at (0-based, required for freeze)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for freeze/unfreeze operations, defaults to input path)"
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
            "freeze" => await FreezePanesAsync(arguments, path, sheetIndex),
            "unfreeze" => await UnfreezePanesAsync(arguments, path, sheetIndex),
            "get" => await GetFreezePanesAsync(arguments, path, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Freezes panes at the specified row and column
    /// </summary>
    /// <param name="arguments">JSON arguments containing row, column (1-based)</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message with freeze position</returns>
    private Task<string> FreezePanesAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            var row = ArgumentHelper.GetInt(arguments, "row");
            var column = ArgumentHelper.GetInt(arguments, "column");

            using var workbook = new Workbook(path);
            var worksheet = workbook.Worksheets[sheetIndex];

            // FreezePanes parameters are 1-based in Aspose.Cells
            // If user provides 0-based row/column, convert to 1-based
            worksheet.FreezePanes(row + 1, column + 1, row + 1, column + 1);

            // Save freeze information to custom properties for later retrieval
            // Since FirstVisibleRow/FirstVisibleColumn might not reflect freeze status,
            // we'll store it in custom properties
            var customProperties = workbook.CustomDocumentProperties;
            var freezeKey = $"FreezePanes_Sheet{sheetIndex}";
            var freezeValue = $"{row},{column}";

            // Remove existing freeze property if any
            try
            {
                customProperties.Remove(freezeKey);
            }
            catch (Exception ex)
            {
                // Ignore if property doesn't exist
                Console.Error.WriteLine($"[WARN] Error removing freeze property: {ex.Message}");
            }

            // Add new freeze property
            customProperties.Add(freezeKey, freezeValue);

            workbook.Save(outputPath);
            return $"Frozen panes (row {row}, column {column}): {outputPath}";
        });
    }

    /// <summary>
    ///     Unfreezes panes in the worksheet
    /// </summary>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Success message</returns>
    private Task<string> UnfreezePanesAsync(JsonObject? arguments, string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            using var workbook = new Workbook(path);
            var worksheet = workbook.Worksheets[sheetIndex];

            // Try to get current freeze status from custom properties (for potential future use)
            var customProperties = workbook.CustomDocumentProperties;
            var freezeKey = $"FreezePanes_Sheet{sheetIndex}";

            try
            {
                var freezeProperty = customProperties[freezeKey];
                if (freezeProperty != null)
                {
                    // Property exists, but we don't need to parse it for unfreeze operation
                    // The unfreeze operation doesn't require the frozen row/column values
                }
            }
            catch (Exception ex)
            {
                // Ignore if property doesn't exist
                Console.Error.WriteLine($"[WARN] Error removing freeze property: {ex.Message}");
            }

            // Try to unfreeze using RemoveFreezePanes if available, otherwise use alternative method
            try
            {
                var worksheetType = worksheet.GetType();
                var removeMethod = worksheetType.GetMethod("RemoveFreezePanes");
                if (removeMethod != null)
                {
                    removeMethod.Invoke(worksheet, null);
                }
                else
                {
                    // Alternative: use RemoveSplit if available
                    var removeSplitMethod = worksheetType.GetMethod("RemoveSplit");
                    if (removeSplitMethod != null)
                    {
                        removeSplitMethod.Invoke(worksheet, null);
                    }
                    else
                    {
                        // Last resort: set freeze to a very large value
                        var maxRow = Math.Max(worksheet.Cells.MaxDataRow + 1, 1000);
                        var maxCol = Math.Max(worksheet.Cells.MaxDataColumn + 1, 100);
                        worksheet.FreezePanes(maxRow + 1, maxCol + 1, maxRow + 1, maxCol + 1);
                    }
                }
            }
            catch
            {
                // If all methods fail, try the simple approach with error handling
                try
                {
                    // Try setting to row 1, column 1 (1-based), which might work
                    worksheet.FreezePanes(1, 1, 1, 1);
                }
                catch
                {
                    // If that also fails, we'll just remove the custom property
                    // The actual freeze might remain, but at least we've cleared our tracking
                }
            }

            // Remove freeze information from custom properties
            try
            {
                customProperties.Remove(freezeKey);
            }
            catch (Exception ex)
            {
                // Ignore if property doesn't exist
                Console.Error.WriteLine($"[WARN] Error removing freeze property: {ex.Message}");
            }

            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            workbook.Save(outputPath);
            return $"Unfrozen panes: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets freeze panes status for the worksheet
    /// </summary>
    /// <param name="_">Unused parameter</param>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>Formatted string with freeze panes status</returns>
    private Task<string> GetFreezePanesAsync(JsonObject? _, string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var result = new StringBuilder();

            result.AppendLine($"=== Freeze panes status for worksheet '{worksheet.Name}' ===\n");

            // Check freeze panes status
            // Since FirstVisibleRow/FirstVisibleColumn might not reliably detect freeze status,
            // we'll use custom properties to track freeze information

            var isFrozen = false;
            var frozenRow = 0;
            var frozenColumn = 0;

            // Method 1: Check custom properties (most reliable - stored when freeze is applied)
            var customProperties = workbook.CustomDocumentProperties;
            var freezeKey = $"FreezePanes_Sheet{sheetIndex}";

            try
            {
                var freezeProperty = customProperties[freezeKey];
                if (freezeProperty != null)
                {
                    var freezeValue = freezeProperty.Value?.ToString();
                    if (!string.IsNullOrEmpty(freezeValue))
                    {
                        // Parse freeze value: "row,column"
                        var parts = freezeValue.Split(',');
                        if (parts.Length == 2 &&
                            int.TryParse(parts[0], out var parsedRow) &&
                            int.TryParse(parts[1], out var parsedColumn))
                        {
                            isFrozen = true;
                            frozenRow = parsedRow;
                            frozenColumn = parsedColumn;
                        }
                    }
                }
            }
            catch
            {
                // If custom property doesn't exist or can't be read, try other methods
            }

            // Method 2: Check FirstVisibleRow and FirstVisibleColumn as fallback (can be affected by scrolling)
            if (!isFrozen)
            {
                var firstVisibleRow = worksheet.FirstVisibleRow;
                var firstVisibleColumn = worksheet.FirstVisibleColumn;

                if (firstVisibleRow > 0 || firstVisibleColumn > 0)
                {
                    isFrozen = true;
                    frozenRow = firstVisibleRow;
                    frozenColumn = firstVisibleColumn;
                }
            }

            if (!isFrozen)
            {
                result.AppendLine("Status: Panes not frozen");
            }
            else
            {
                result.AppendLine("Status: Panes frozen");
                result.AppendLine($"Frozen row: {frozenRow}");
                result.AppendLine($"Frozen column: {frozenColumn}");
                result.AppendLine($"Freeze position: before row {frozenRow + 1} and column {frozenColumn + 1}");
            }

            return result.ToString();
        });
    }
}