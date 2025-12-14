using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing Excel named ranges (add, delete, get)
/// Merges: ExcelAddNamedRangeTool, ExcelDeleteNamedRangeTool, ExcelGetNamedRangesTool
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
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation", "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);

        return operation.ToLower() switch
        {
            "add" => await AddNamedRangeAsync(arguments, path),
            "delete" => await DeleteNamedRangeAsync(arguments, path),
            "get" => await GetNamedRangesAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    /// Adds a named range to the workbook
    /// </summary>
    /// <param name="arguments">JSON arguments containing name and range</param>
    /// <param name="path">Excel file path</param>
    /// <returns>Success message with named range details</returns>
    private async Task<string> AddNamedRangeAsync(JsonObject? arguments, string path)
    {
        var name = ArgumentHelper.GetString(arguments, "name", "name");
        var range = ArgumentHelper.GetString(arguments, "range", "range");
        var comment = arguments?["comment"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        var names = workbook.Worksheets.Names;
        
        // Convert range to proper refersTo format (e.g., "Sheet1!A1:A3")
        var worksheet = workbook.Worksheets[0]; // Use first worksheet as default
        string refersTo;
        
        // Check if range already contains sheet reference (e.g., "Sheet1!A1:A3")
        if (range.Contains("!"))
        {
            refersTo = range;
        }
        else
        {
            // Add sheet reference if not present
            // Escape single quotes in sheet name if present
            var sheetName = worksheet.Name.Replace("'", "''");
            refersTo = $"'{sheetName}'!{range}";
        }
        
        // Use the correct API: Add(name) first, then set RefersTo
        // According to Aspose.Cells documentation, this is the correct way
        try
        {
            // Check if name already exists using get operation's method
            using var checkWorkbook = new Workbook(path);
            var checkNames = checkWorkbook.Worksheets.Names;
            for (int i = 0; i < checkNames.Count; i++)
            {
                try
                {
                    var checkName = checkNames[i];
                    var checkText = checkName.Text;
                    if (!string.IsNullOrEmpty(checkText) && checkText == name)
                    {
                        throw new ArgumentException($"名稱範圍 '{name}' 已存在");
                    }
                }
                catch (ArgumentException)
                {
                    throw;
                }
                catch
                {
                    continue;
                }
            }
            
            // Add the name first (without refersTo)
            var nameIndex = names.Add(name);
            var namedRange = names[nameIndex];
            
            // Then set RefersTo property
            namedRange.RefersTo = refersTo;
            
            // Set comment if provided
            if (!string.IsNullOrEmpty(comment))
            {
                namedRange.Comment = comment;
            }
            
            // Save the workbook to persist the changes
            workbook.Save(path);
            
            return await Task.FromResult($"成功添加名稱範圍 '{name}'\n引用: {refersTo}\n輸出: {path}");
        }
        catch (ArgumentException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"無法創建名稱範圍 '{name}'，引用: {refersTo}。錯誤: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Deletes a named range from the workbook
    /// </summary>
    /// <param name="arguments">JSON arguments containing name</param>
    /// <param name="path">Excel file path</param>
    /// <returns>Success message</returns>
    private async Task<string> DeleteNamedRangeAsync(JsonObject? arguments, string path)
    {
        var name = ArgumentHelper.GetString(arguments, "name", "name");

        using var workbook = new Workbook(path);
        var names = workbook.Worksheets.Names;
        
        Name? namedRange = null;
        try
        {
            namedRange = names[name];
        }
        catch
        {
            throw new ArgumentException($"名稱範圍 '{name}' 不存在");
        }
        
        if (namedRange == null)
        {
            throw new ArgumentException($"名稱範圍 '{name}' 不存在");
        }
        
        var refersTo = namedRange.RefersTo;
        // Find the index of the named range
        int indexToRemove = -1;
        for (int i = 0; i < names.Count; i++)
        {
            if (names[i] == namedRange)
            {
                indexToRemove = i;
                break;
            }
        }
        
        if (indexToRemove >= 0)
        {
            names.RemoveAt(indexToRemove);
        }
        workbook.Save(path);
        
        var remainingCount = names.Count;
        
        return await Task.FromResult($"成功刪除名稱範圍 '{name}'\n原引用: {refersTo}\n工作簿剩餘名稱範圍數: {remainingCount}\n輸出: {path}");
    }

    /// <summary>
    /// Gets all named ranges from the workbook
    /// </summary>
    /// <param name="arguments">JSON arguments (no specific parameters required)</param>
    /// <param name="path">Excel file path</param>
    /// <returns>Formatted string with all named ranges</returns>
    private async Task<string> GetNamedRangesAsync(JsonObject? arguments, string path)
    {
        using var workbook = new Workbook(path);
        var names = workbook.Worksheets.Names;
        var result = new StringBuilder();

        result.AppendLine("=== Excel 工作簿的名稱範圍資訊 ===\n");
        result.AppendLine($"總名稱範圍數: {names.Count}\n");

        if (names.Count == 0)
        {
            result.AppendLine("未找到名稱範圍");
            return await Task.FromResult(result.ToString());
        }

        for (int i = 0; i < names.Count; i++)
        {
            var name = names[i];
            result.AppendLine($"【名稱範圍 {i}】");
            result.AppendLine($"名稱: {name.Text}");
            result.AppendLine($"引用: {name.RefersTo}");
            result.AppendLine($"註解: {name.Comment ?? "(無)"}");
            result.AppendLine($"是否可見: {name.IsVisible}");
            result.AppendLine();
        }

        return await Task.FromResult(result.ToString());
    }
}

