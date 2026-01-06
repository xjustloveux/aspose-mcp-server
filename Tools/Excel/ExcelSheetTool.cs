using System.ComponentModel;
using System.Text.Json;
using Aspose.Cells;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel sheets (add, delete, get, rename, move, copy, hide)
/// </summary>
[McpServerToolType]
public class ExcelSheetTool
{
    /// <summary>
    ///     Characters that are not allowed in Excel sheet names.
    /// </summary>
    private static readonly char[] InvalidSheetNameChars = ['\\', '/', '?', '*', '[', ']', ':'];

    /// <summary>
    ///     Session identity accessor for session isolation support.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Document session manager for in-memory editing support.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelSheetTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    /// <param name="identityAccessor">Optional session identity accessor for session isolation.</param>
    public ExcelSheetTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes an Excel sheet operation (add, delete, get, rename, move, copy, or hide).
    /// </summary>
    /// <param name="operation">The operation to perform: add, delete, get, rename, move, copy, or hide.</param>
    /// <param name="path">Excel file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="sheetIndex">Sheet index (0-based, required for delete/rename/move/copy/hide).</param>
    /// <param name="sheetName">Name of the sheet (required for add operation).</param>
    /// <param name="newName">New name for the sheet (required for rename, max 31 characters).</param>
    /// <param name="insertAt">Position to insert the sheet (0-based, optional for add/move).</param>
    /// <param name="targetIndex">Target index for move/copy operation (0-based).</param>
    /// <param name="copyToPath">Target file path for copy operation (optional).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    /// <exception cref="InvalidOperationException">Thrown when attempting to delete the last worksheet.</exception>
    [McpServerTool(Name = "excel_sheet")]
    [Description(@"Manage Excel sheets. Supports 7 operations: add, delete, get, rename, move, copy, hide.

Usage examples:
- Add sheet: excel_sheet(operation='add', path='book.xlsx', sheetName='New Sheet')
- Delete sheet: excel_sheet(operation='delete', path='book.xlsx', sheetIndex=1)
- Get sheets: excel_sheet(operation='get', path='book.xlsx')
- Rename sheet: excel_sheet(operation='rename', path='book.xlsx', sheetIndex=0, newName='Renamed')
- Move sheet: excel_sheet(operation='move', path='book.xlsx', sheetIndex=0, insertAt=2)
- Copy sheet: excel_sheet(operation='copy', path='book.xlsx', sheetIndex=0, newName='Copy')
- Hide sheet: excel_sheet(operation='hide', path='book.xlsx', sheetIndex=1)")]
    public string Execute(
        [Description("Operation to perform: add, delete, get, rename, move, copy, hide")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, required for delete/rename/move/copy/hide)")]
        int sheetIndex = 0,
        [Description("Name of the sheet (required for add operation)")]
        string? sheetName = null,
        [Description("New name for the sheet (required for rename, max 31 characters)")]
        string? newName = null,
        [Description("Position to insert the sheet (0-based, optional for add/move)")]
        int? insertAt = null,
        [Description("Target index for move/copy operation (0-based)")]
        int? targetIndex = null,
        [Description("Target file path for copy operation (optional)")]
        string? copyToPath = null)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLowerInvariant() switch
        {
            "add" => AddSheet(ctx, outputPath, sheetName, insertAt),
            "delete" => DeleteSheet(ctx, outputPath, sheetIndex),
            "get" => GetSheets(ctx),
            "rename" => RenameSheet(ctx, outputPath, sheetIndex, newName),
            "move" => MoveSheet(ctx, outputPath, sheetIndex, targetIndex, insertAt),
            "copy" => CopySheet(ctx, outputPath, sheetIndex, targetIndex, copyToPath),
            "hide" => HideSheet(ctx, outputPath, sheetIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Validates that the sheet name meets Excel requirements.
    /// </summary>
    /// <param name="name">The sheet name to validate.</param>
    /// <param name="paramName">The parameter name for error messages.</param>
    /// <exception cref="ArgumentException">
    ///     Thrown when the sheet name is empty, exceeds 31 characters, or contains invalid
    ///     characters.
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
    ///     Adds a new worksheet to the workbook.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetName">The name of the new worksheet.</param>
    /// <param name="insertAt">The position to insert the worksheet at.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when sheetName is empty, invalid, already exists, or insertAt is out of
    ///     range.
    /// </exception>
    private static string AddSheet(DocumentContext<Workbook> ctx, string? outputPath, string? sheetName, int? insertAt)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException("sheetName is required for add operation");

        sheetName = sheetName.Trim();
        ValidateSheetName(sheetName, "sheetName");

        var workbook = ctx.Document;

        var duplicate =
            workbook.Worksheets.Any(ws => string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase));
        if (duplicate)
            throw new ArgumentException($"Worksheet name '{sheetName}' already exists in the workbook");

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

        ctx.Save(outputPath);
        return $"Worksheet '{sheetName}' added. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Deletes a worksheet from the workbook.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The index of the worksheet to delete.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when the sheet index is out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when attempting to delete the last worksheet.</exception>
    private static string DeleteSheet(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex)
    {
        var workbook = ctx.Document;
        ExcelHelper.ValidateSheetIndex(sheetIndex, workbook);

        if (workbook.Worksheets.Count <= 1)
            throw new InvalidOperationException("Cannot delete the last worksheet");

        var sheetName = workbook.Worksheets[sheetIndex].Name;
        workbook.Worksheets.RemoveAt(sheetIndex);

        ctx.Save(outputPath);
        return $"Worksheet '{sheetName}' (index {sheetIndex}) deleted. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Gets information about all worksheets in the workbook.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <returns>A JSON string containing information about all worksheets.</returns>
    private static string GetSheets(DocumentContext<Workbook> ctx)
    {
        var workbook = ctx.Document;

        if (workbook.Worksheets.Count == 0)
        {
            var emptyResult = new
            {
                count = 0,
                workbookName = ctx.SourcePath != null ? Path.GetFileName(ctx.SourcePath) : "session",
                items = Array.Empty<object>(),
                message = "No worksheets found"
            };
            return JsonSerializer.Serialize(emptyResult, new JsonSerializerOptions { WriteIndented = true });
        }

        List<object> sheetList = [];
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
            workbookName = ctx.SourcePath != null ? Path.GetFileName(ctx.SourcePath) : "session",
            items = sheetList
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Renames a worksheet.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The index of the worksheet to rename.</param>
    /// <param name="newName">The new name for the worksheet.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when newName is empty, invalid, or already exists.</exception>
    private static string RenameSheet(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        string? newName)
    {
        if (string.IsNullOrEmpty(newName))
            throw new ArgumentException("newName is required for rename operation");

        newName = newName.Trim();
        ValidateSheetName(newName, "newName");

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var oldName = worksheet.Name;

        var duplicate = workbook.Worksheets.Any(ws =>
            ws != worksheet && string.Equals(ws.Name, newName, StringComparison.OrdinalIgnoreCase));
        if (duplicate)
            throw new ArgumentException($"Worksheet name '{newName}' already exists in the workbook");

        worksheet.Name = newName;

        ctx.Save(outputPath);
        return $"Worksheet '{oldName}' renamed to '{newName}'. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Moves a worksheet to a different position.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The index of the worksheet to move.</param>
    /// <param name="targetIndex">The target position index.</param>
    /// <param name="insertAt">Alternative target position index.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when neither targetIndex nor insertAt is provided, or when indices are out
    ///     of range.
    /// </exception>
    private static string MoveSheet(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex, int? targetIndex,
        int? insertAt)
    {
        if (!targetIndex.HasValue && !insertAt.HasValue)
            throw new ArgumentException("Either targetIndex or insertAt is required for move operation");

        var finalTargetIndex = targetIndex ?? insertAt!.Value;

        var workbook = ctx.Document;

        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
            throw new ArgumentException(
                $"Worksheet index {sheetIndex} is out of range (workbook has {workbook.Worksheets.Count} worksheets)");

        if (finalTargetIndex < 0 || finalTargetIndex >= workbook.Worksheets.Count)
            throw new ArgumentException(
                $"Target index {finalTargetIndex} is out of range (workbook has {workbook.Worksheets.Count} worksheets)");

        if (sheetIndex == finalTargetIndex)
            return $"Worksheet is already at position {sheetIndex}, no move needed. {ctx.GetOutputMessage(outputPath)}";

        var sheetName = workbook.Worksheets[sheetIndex].Name;
        var worksheet = workbook.Worksheets[sheetIndex];

        worksheet.MoveTo(finalTargetIndex);

        ctx.Save(outputPath);
        return
            $"Worksheet '{sheetName}' moved from position {sheetIndex} to {finalTargetIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Copies a worksheet within the same workbook or to another file.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The index of the worksheet to copy.</param>
    /// <param name="targetIndex">The target position index for the copied worksheet.</param>
    /// <param name="copyToPath">The path to copy the worksheet to an external file.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when sheetIndex or targetIndex is out of range.</exception>
    private static string CopySheet(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex, int? targetIndex,
        string? copyToPath)
    {
        if (!string.IsNullOrEmpty(copyToPath))
            SecurityHelper.ValidateFilePath(copyToPath, "copyToPath", true);

        var workbook = ctx.Document;

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

        ctx.Save(outputPath);
        return $"Worksheet '{sheetName}' copied to position {targetIndex.Value}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Toggles the visibility of a worksheet.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The index of the worksheet to hide or show.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    private static string HideSheet(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var sheetName = worksheet.Name;

        if (worksheet.IsVisible)
        {
            worksheet.IsVisible = false;
            ctx.Save(outputPath);
            return $"Worksheet '{sheetName}' hidden. {ctx.GetOutputMessage(outputPath)}";
        }

        worksheet.IsVisible = true;
        ctx.Save(outputPath);
        return $"Worksheet '{sheetName}' shown. {ctx.GetOutputMessage(outputPath)}";
    }
}