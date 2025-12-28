using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel sheets (add, delete, get, rename, move, copy, hide)
/// </summary>
public class ExcelSheetTool : IAsposeTool
{
    private const string OperationAdd = "add";
    private const string OperationDelete = "delete";
    private const string OperationGet = "get";
    private const string OperationRename = "rename";
    private const string OperationMove = "move";
    private const string OperationCopy = "copy";
    private const string OperationHide = "hide";

    private static readonly char[] InvalidSheetNameChars = ['\\', '/', '?', '*', '[', ']', ':'];

    /// <summary>
    ///     Gets the description of the tool and its usage examples.
    /// </summary>
    public string Description =>
        @"Manage Excel sheets. Supports 7 operations: add, delete, get, rename, move, copy, hide.

Usage examples:
- Add sheet: excel_sheet(operation='add', path='book.xlsx', sheetName='New Sheet')
- Delete sheet: excel_sheet(operation='delete', path='book.xlsx', sheetIndex=1)
- Get sheets: excel_sheet(operation='get', path='book.xlsx')
- Rename sheet: excel_sheet(operation='rename', path='book.xlsx', sheetIndex=0, newName='Renamed')
- Move sheet: excel_sheet(operation='move', path='book.xlsx', sheetIndex=0, insertAt=2)
- Copy sheet: excel_sheet(operation='copy', path='book.xlsx', sheetIndex=0, newName='Copy')
- Hide sheet: excel_sheet(operation='hide', path='book.xlsx', sheetIndex=1)";

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
- 'add': Add a new sheet (required params: path, sheetName)
- 'delete': Delete a sheet (required params: path, sheetIndex)
- 'get': Get all sheets (required params: path)
- 'rename': Rename a sheet (required params: path, sheetIndex, newName)
- 'move': Move a sheet (required params: path, sheetIndex, insertAt)
- 'copy': Copy a sheet (required params: path, sheetIndex, newName)
- 'hide': Hide a sheet (required params: path, sheetIndex)",
                @enum = new[] { "add", "delete", "get", "rename", "move", "copy", "hide" }
            },
            path = new
            {
                type = "string",
                description = "Excel file path (required for all operations)"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, required for delete/rename/move/copy/hide)"
            },
            sheetName = new
            {
                type = "string",
                description = "Name of the sheet (required for add operation)"
            },
            newName = new
            {
                type = "string",
                description =
                    "New name for the sheet (required for rename, maximum 31 characters recommended for Excel compatibility)"
            },
            insertAt = new
            {
                type = "number",
                description =
                    "Position to insert the sheet (0-based, optional for add, optional for move as alternative to targetIndex)"
            },
            targetIndex = new
            {
                type = "number",
                description =
                    "Target index for move/copy operation (0-based, required for move, or use insertAt as alternative)"
            },
            copyToPath = new
            {
                type = "string",
                description =
                    "Target file path for copy operation (optional, if not provided copies within same workbook)"
            },
            outputPath = new
            {
                type = "string",
                description =
                    "Output file path (optional, for add/delete/rename/move/copy/hide operations, defaults to input path)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid</exception>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        return operation.ToLowerInvariant() switch
        {
            OperationAdd => await AddSheetAsync(path, outputPath, arguments),
            OperationDelete => await DeleteSheetAsync(path, outputPath, arguments),
            OperationGet => await GetSheetsAsync(path),
            OperationRename => await RenameSheetAsync(path, outputPath, arguments),
            OperationMove => await MoveSheetAsync(path, outputPath, arguments),
            OperationCopy => await CopySheetAsync(path, outputPath, arguments),
            OperationHide => await HideSheetAsync(path, outputPath, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }


    /// <summary>
    ///     Validates a sheet name according to Excel's naming rules
    /// </summary>
    /// <param name="name">The sheet name to validate</param>
    /// <param name="paramName">The parameter name for error messages</param>
    /// <exception cref="ArgumentException">
    ///     Thrown when the name is empty, exceeds 31 characters, or contains invalid characters (\ / ? * [ ] :)
    /// </exception>
    private static void ValidateSheetName(string name, string paramName)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException($"{paramName} cannot be empty");

        if (name.Length > 31)
            throw new ArgumentException(
                $"{paramName} '{name}' (length: {name.Length}) exceeds Excel's limit of 31 characters");

        var invalidCharIndex = name.IndexOfAny(InvalidSheetNameChars);
        if (invalidCharIndex >= 0)
            throw new ArgumentException(
                $"{paramName} contains invalid character '{name[invalidCharIndex]}'. Sheet names cannot contain: \\ / ? * [ ] :");
    }

    /// <summary>
    ///     Adds a new worksheet to the workbook
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing sheetName and optional insertAt</param>
    /// <returns>Success message with worksheet name</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when sheetName is invalid, duplicated, or insertAt is out of range
    /// </exception>
    private Task<string> AddSheetAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var sheetName = ArgumentHelper.GetString(arguments, "sheetName").Trim();
            var insertAt = ArgumentHelper.GetIntNullable(arguments, "insertAt");

            ValidateSheetName(sheetName, "sheetName");

            using var workbook = new Workbook(path);

            var duplicate =
                workbook.Worksheets.Any(ws => string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase));
            if (duplicate) throw new ArgumentException($"Worksheet name '{sheetName}' already exists in the workbook");

            Worksheet newSheet;
            if (insertAt.HasValue)
            {
                if (insertAt.Value < 0 || insertAt.Value > workbook.Worksheets.Count)
                    throw new ArgumentException($"insertAt must be between 0 and {workbook.Worksheets.Count}");

                if (insertAt.Value == workbook.Worksheets.Count)
                {
                    var addedIndex = workbook.Worksheets.Add();
                    newSheet = workbook.Worksheets[addedIndex];
                }
                else
                {
                    workbook.Worksheets.Insert(insertAt.Value, SheetType.Worksheet);
                    newSheet = workbook.Worksheets[insertAt.Value];
                }
            }
            else
            {
                var addedIndex = workbook.Worksheets.Add();
                newSheet = workbook.Worksheets[addedIndex];
            }

            newSheet.Name = sheetName;
            workbook.Save(outputPath);

            return $"Worksheet '{sheetName}' added. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes a worksheet from the workbook
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing sheetIndex</param>
    /// <returns>Success message with deleted sheet name</returns>
    /// <exception cref="ArgumentException">Thrown when sheetIndex is out of range</exception>
    /// <exception cref="InvalidOperationException">Thrown when trying to delete the last worksheet</exception>
    private Task<string> DeleteSheetAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex");

            using var workbook = new Workbook(path);
            ExcelHelper.ValidateSheetIndex(sheetIndex, workbook);

            if (workbook.Worksheets.Count <= 1) throw new InvalidOperationException("Cannot delete the last worksheet");

            var sheetName = workbook.Worksheets[sheetIndex].Name;
            workbook.Worksheets.RemoveAt(sheetIndex);
            workbook.Save(outputPath);

            return $"Worksheet '{sheetName}' (index {sheetIndex}) deleted. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets information about all worksheets in the workbook
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <returns>JSON string with worksheet list</returns>
    private Task<string> GetSheetsAsync(string path)
    {
        return Task.Run(() =>
        {
            using var workbook = new Workbook(path);

            if (workbook.Worksheets.Count == 0)
            {
                var emptyResult = new
                {
                    count = 0,
                    workbookName = Path.GetFileName(path),
                    items = Array.Empty<object>(),
                    message = "No worksheets found"
                };
                return JsonSerializer.Serialize(emptyResult, new JsonSerializerOptions { WriteIndented = true });
            }

            var sheetList = new List<object>();
            for (var i = 0; i < workbook.Worksheets.Count; i++)
            {
                var worksheet = workbook.Worksheets[i];
                sheetList.Add(new
                {
                    index = i,
                    name = worksheet.Name,
                    visibility = worksheet.IsVisible ? "Visible" : "Hidden"
                });
            }

            var result = new
            {
                count = workbook.Worksheets.Count,
                workbookName = Path.GetFileName(path),
                items = sheetList
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }

    /// <summary>
    ///     Renames a worksheet
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing sheetIndex and newName</param>
    /// <returns>Success message with old and new names</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when sheetIndex is out of range, newName is invalid, or duplicated
    /// </exception>
    private Task<string> RenameSheetAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex");
            var newName = ArgumentHelper.GetString(arguments, "newName").Trim();

            ValidateSheetName(newName, "newName");

            using var workbook = new Workbook(path);

            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var oldName = worksheet.Name;

            var duplicate = workbook.Worksheets.Any(ws =>
                ws != worksheet && string.Equals(ws.Name, newName, StringComparison.OrdinalIgnoreCase));
            if (duplicate) throw new ArgumentException($"Worksheet name '{newName}' already exists in the workbook");


            worksheet.Name = newName;
            workbook.Save(outputPath);

            return $"Worksheet '{oldName}' renamed to '{newName}'. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Moves a worksheet to a different position
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing sheetIndex and targetIndex or insertAt</param>
    /// <returns>Success message with move details</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when sheetIndex or targetIndex/insertAt is out of range, or neither is provided
    /// </exception>
    private Task<string> MoveSheetAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex");
            var targetIndex = ArgumentHelper.GetIntNullable(arguments, "targetIndex");
            var insertAt = ArgumentHelper.GetIntNullable(arguments, "insertAt");

            if (!targetIndex.HasValue && !insertAt.HasValue)
                throw new ArgumentException("Either targetIndex or insertAt is required for move operation");

            var finalTargetIndex = targetIndex ?? insertAt!.Value;

            using var workbook = new Workbook(path);

            if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
                throw new ArgumentException(
                    $"Worksheet index {sheetIndex} is out of range (workbook has {workbook.Worksheets.Count} worksheets)");

            if (finalTargetIndex < 0 || finalTargetIndex >= workbook.Worksheets.Count)
                throw new ArgumentException(
                    $"Target index {finalTargetIndex} is out of range (workbook has {workbook.Worksheets.Count} worksheets)");

            if (sheetIndex == finalTargetIndex)
                return $"Worksheet is already at position {sheetIndex}, no move needed. Output: {path}";

            var sheetName = workbook.Worksheets[sheetIndex].Name;
            var worksheet = workbook.Worksheets[sheetIndex];

            worksheet.MoveTo(finalTargetIndex);
            workbook.Save(outputPath);

            return
                $"Worksheet '{sheetName}' moved from position {sheetIndex} to {finalTargetIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Copies a worksheet within the same workbook or to an external file
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing sheetIndex, optional targetIndex or copyToPath</param>
    /// <returns>Success message with copy details</returns>
    /// <exception cref="ArgumentException">Thrown when sheetIndex or targetIndex is out of range</exception>
    private Task<string> CopySheetAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex");
            var targetIndex = ArgumentHelper.GetIntNullable(arguments, "targetIndex");
            var copyToPath = ArgumentHelper.GetStringNullable(arguments, "copyToPath");
            if (!string.IsNullOrEmpty(copyToPath)) SecurityHelper.ValidateFilePath(copyToPath, "copyToPath", true);

            using var workbook = new Workbook(path);

            if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
                throw new ArgumentException(
                    $"Worksheet index {sheetIndex} is out of range (workbook has {workbook.Worksheets.Count} worksheets)");

            var sourceSheet = workbook.Worksheets[sheetIndex];
            var sheetName = sourceSheet.Name;

            if (!string.IsNullOrEmpty(copyToPath))
            {
                using var targetWorkbook = new Workbook();

                targetWorkbook.Worksheets[0].Copy(sourceSheet);
                targetWorkbook.Worksheets[0].Name = sheetName;
                targetWorkbook.Save(copyToPath);
                return $"Worksheet '{sheetName}' copied to external file. Output: {copyToPath}";
            }

            targetIndex ??= workbook.Worksheets.Count;

            if (targetIndex.Value < 0 || targetIndex.Value > workbook.Worksheets.Count)
                throw new ArgumentException(
                    $"Target index {targetIndex.Value} is out of range (workbook has {workbook.Worksheets.Count} worksheets)");

            _ = workbook.Worksheets.AddCopy(sheetIndex);
            workbook.Save(outputPath);
            return $"Worksheet '{sheetName}' copied to position {targetIndex.Value}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Toggles the visibility of a worksheet (hides if visible, shows if hidden)
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing sheetIndex</param>
    /// <returns>Success message with visibility status</returns>
    /// <exception cref="ArgumentException">Thrown when sheetIndex is out of range</exception>
    private Task<string> HideSheetAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex");

            using var workbook = new Workbook(path);

            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var sheetName = worksheet.Name;

            if (worksheet.IsVisible)
            {
                worksheet.IsVisible = false;
                workbook.Save(outputPath);
                return $"Worksheet '{sheetName}' hidden. Output: {outputPath}";
            }

            worksheet.IsVisible = true;
            workbook.Save(outputPath);
            return $"Worksheet '{sheetName}' shown. Output: {outputPath}";
        });
    }
}