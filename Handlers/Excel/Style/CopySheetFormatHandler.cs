using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Style;

/// <summary>
///     Handler for copying format from one sheet to another in Excel workbooks.
/// </summary>
public class CopySheetFormatHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "copy_sheet_format";

    /// <summary>
    ///     Copies format from source sheet to target sheet.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: sourceSheetIndex, targetSheetIndex
    ///     Optional: copyColumnWidths, copyRowHeights
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sourceSheetIndex = parameters.GetOptional("sourceSheetIndex", 0);
        var targetSheetIndex = parameters.GetOptional("targetSheetIndex", 0);
        var copyColumnWidths = parameters.GetOptional("copyColumnWidths", true);
        var copyRowHeights = parameters.GetOptional("copyRowHeights", true);

        var workbook = context.Document;
        var sourceSheet = ExcelHelper.GetWorksheet(workbook, sourceSheetIndex);
        var targetSheet = ExcelHelper.GetWorksheet(workbook, targetSheetIndex);

        if (copyColumnWidths)
        {
            var maxCol = Math.Max(sourceSheet.Cells.MaxColumn, sourceSheet.Cells.MaxDataColumn);
            maxCol = Math.Max(maxCol, 0);
            for (var i = 0; i <= maxCol; i++)
                targetSheet.Cells.SetColumnWidth(i, sourceSheet.Cells.GetColumnWidth(i));
        }

        if (copyRowHeights)
        {
            var maxRow = Math.Max(sourceSheet.Cells.MaxRow, sourceSheet.Cells.MaxDataRow);
            maxRow = Math.Max(maxRow, 0);
            for (var i = 0; i <= maxRow; i++)
                targetSheet.Cells.SetRowHeight(i, sourceSheet.Cells.GetRowHeight(i));
        }

        MarkModified(context);

        return Success($"Sheet format copied from sheet {sourceSheetIndex} to sheet {targetSheetIndex}.");
    }
}
