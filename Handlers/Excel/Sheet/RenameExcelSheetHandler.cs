using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Sheet;

/// <summary>
///     Handler for renaming worksheets in Excel workbooks.
/// </summary>
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
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetRequired<int>("sheetIndex");
        var newName = parameters.GetRequired<string>("newName");

        newName = newName.Trim();
        ExcelSheetHelper.ValidateSheetName(newName, "newName");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var oldName = worksheet.Name;

        var duplicate = workbook.Worksheets.Any(ws =>
            ws != worksheet && string.Equals(ws.Name, newName, StringComparison.OrdinalIgnoreCase));
        if (duplicate)
            throw new ArgumentException($"Worksheet name '{newName}' already exists in the workbook");

        worksheet.Name = newName;

        MarkModified(context);

        return Success($"Worksheet '{oldName}' renamed to '{newName}'.");
    }
}
