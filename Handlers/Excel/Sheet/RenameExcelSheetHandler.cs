using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Sheet;

/// <summary>
///     Handler for renaming worksheets in Excel workbooks.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class RenameExcelSheetHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "rename";

    /// <summary>
    ///     Renames a worksheet in the workbook.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: sheetIndex (0-based), newName (max 31 characters)
    /// </param>
    /// <returns>Success message with operation details.</returns>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractRenameExcelSheetParameters(parameters);

        var newName = p.NewName.Trim();
        ExcelSheetHelper.ValidateSheetName(newName, "newName");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);
        var oldName = worksheet.Name;

        var duplicate = workbook.Worksheets.Any(ws =>
            ws != worksheet && string.Equals(ws.Name, newName, StringComparison.OrdinalIgnoreCase));
        if (duplicate)
            throw new ArgumentException($"Worksheet name '{newName}' already exists in the workbook");

        worksheet.Name = newName;

        MarkModified(context);

        return new SuccessResult { Message = $"Worksheet '{oldName}' renamed to '{newName}'." };
    }

    private static RenameExcelSheetParameters ExtractRenameExcelSheetParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetRequired<int>("sheetIndex");
        var newName = parameters.GetRequired<string>("newName");

        return new RenameExcelSheetParameters(sheetIndex, newName);
    }

    private sealed record RenameExcelSheetParameters(int SheetIndex, string NewName);
}
