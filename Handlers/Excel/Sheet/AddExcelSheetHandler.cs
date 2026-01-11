using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.Sheet;

/// <summary>
///     Handler for adding worksheets to Excel workbooks.
/// </summary>
public class AddExcelSheetHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a new worksheet to the workbook.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: sheetName
    ///     Optional: insertAt (position to insert, 0-based)
    /// </param>
    /// <returns>Success message with operation details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetName = parameters.GetRequired<string>("sheetName");
        var insertAt = parameters.GetOptional<int?>("insertAt");

        sheetName = sheetName.Trim();
        ExcelSheetHelper.ValidateSheetName(sheetName, "sheetName");

        var workbook = context.Document;

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

        MarkModified(context);

        return Success($"Worksheet '{sheetName}' added.");
    }
}
