using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Sheet;

/// <summary>
///     Handler for copying worksheets in Excel workbooks.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class CopyExcelSheetHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "copy";

    /// <summary>
    ///     Copies a worksheet within the same workbook or to an external file.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: sheetIndex (0-based index of sheet to copy)
    ///     Optional: targetIndex (position for copy), copyToPath (external file path), newName (name for copied sheet)
    /// </param>
    /// <returns>Success message with operation details.</returns>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractCopyExcelSheetParameters(parameters);

        if (!string.IsNullOrEmpty(p.CopyToPath))
            SecurityHelper.ValidateFilePath(p.CopyToPath, "copyToPath", true);

        var newName = p.NewName?.Trim();
        if (!string.IsNullOrEmpty(newName))
            ExcelSheetHelper.ValidateSheetName(newName, "newName");

        var workbook = context.Document;

        if (p.SheetIndex < 0 || p.SheetIndex >= workbook.Worksheets.Count)
            throw new ArgumentException(
                $"Worksheet index {p.SheetIndex} is out of range (workbook has {workbook.Worksheets.Count} worksheets)");

        var sourceSheet = workbook.Worksheets[p.SheetIndex];
        var sheetName = sourceSheet.Name;

        if (!string.IsNullOrEmpty(p.CopyToPath))
        {
            using var targetWorkbook = new Workbook();
            targetWorkbook.Worksheets[0].Copy(sourceSheet);
            targetWorkbook.Worksheets[0].Name = newName ?? sheetName;
            targetWorkbook.Save(p.CopyToPath);
            return new SuccessResult
                { Message = $"Worksheet '{sheetName}' copied to external file. Output: {p.CopyToPath}" };
        }

        if (!string.IsNullOrEmpty(newName))
        {
            var duplicate = workbook.Worksheets.Any(ws =>
                string.Equals(ws.Name, newName, StringComparison.OrdinalIgnoreCase));
            if (duplicate)
                throw new ArgumentException($"Worksheet name '{newName}' already exists in the workbook");
        }

        var targetIndex = p.TargetIndex ?? workbook.Worksheets.Count;

        if (targetIndex < 0 || targetIndex > workbook.Worksheets.Count)
            throw new ArgumentException(
                $"Target index {targetIndex} is out of range (workbook has {workbook.Worksheets.Count} worksheets)");

        var copyIndex = workbook.Worksheets.AddCopy(p.SheetIndex);

        if (!string.IsNullOrEmpty(newName))
            workbook.Worksheets[copyIndex].Name = newName;

        MarkModified(context);

        return new SuccessResult { Message = $"Worksheet '{sheetName}' copied to position {targetIndex}." };
    }

    private static CopyExcelSheetParameters ExtractCopyExcelSheetParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetRequired<int>("sheetIndex");
        var targetIndex = parameters.GetOptional<int?>("targetIndex");
        var copyToPath = parameters.GetOptional<string?>("copyToPath");
        var newName = parameters.GetOptional<string?>("newName");

        return new CopyExcelSheetParameters(sheetIndex, targetIndex, copyToPath, newName);
    }

    private sealed record CopyExcelSheetParameters(
        int SheetIndex,
        int? TargetIndex,
        string? CopyToPath,
        string? NewName);
}
