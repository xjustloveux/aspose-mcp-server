using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Sheet;

/// <summary>
///     Handler for copying worksheets in Excel workbooks.
/// </summary>
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
    ///     Optional: targetIndex (position for copy), copyToPath (external file path)
    /// </param>
    /// <returns>Success message with operation details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetRequired<int>("sheetIndex");
        var targetIndex = parameters.GetOptional<int?>("targetIndex");
        var copyToPath = parameters.GetOptional<string?>("copyToPath");

        if (!string.IsNullOrEmpty(copyToPath))
            SecurityHelper.ValidateFilePath(copyToPath, "copyToPath", true);

        var workbook = context.Document;

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
            return Success($"Worksheet '{sheetName}' copied to external file. Output: {copyToPath}");
        }

        targetIndex ??= workbook.Worksheets.Count;

        if (targetIndex.Value < 0 || targetIndex.Value > workbook.Worksheets.Count)
            throw new ArgumentException(
                $"Target index {targetIndex.Value} is out of range (workbook has {workbook.Worksheets.Count} worksheets)");

        _ = workbook.Worksheets.AddCopy(sheetIndex);

        MarkModified(context);

        return Success($"Worksheet '{sheetName}' copied to position {targetIndex.Value}.");
    }
}
