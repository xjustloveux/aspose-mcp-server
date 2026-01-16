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
        var p = ExtractCopySheetFormatParameters(parameters);

        var workbook = context.Document;
        var sourceSheet = ExcelHelper.GetWorksheet(workbook, p.SourceSheetIndex);
        var targetSheet = ExcelHelper.GetWorksheet(workbook, p.TargetSheetIndex);

        if (p.CopyColumnWidths)
        {
            var maxCol = Math.Max(sourceSheet.Cells.MaxColumn, sourceSheet.Cells.MaxDataColumn);
            maxCol = Math.Max(maxCol, 0);
            for (var i = 0; i <= maxCol; i++)
                targetSheet.Cells.SetColumnWidth(i, sourceSheet.Cells.GetColumnWidth(i));
        }

        if (p.CopyRowHeights)
        {
            var maxRow = Math.Max(sourceSheet.Cells.MaxRow, sourceSheet.Cells.MaxDataRow);
            maxRow = Math.Max(maxRow, 0);
            for (var i = 0; i <= maxRow; i++)
                targetSheet.Cells.SetRowHeight(i, sourceSheet.Cells.GetRowHeight(i));
        }

        MarkModified(context);

        return Success($"Sheet format copied from sheet {p.SourceSheetIndex} to sheet {p.TargetSheetIndex}.");
    }

    /// <summary>
    ///     Extracts copy sheet format parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>A CopySheetFormatParameters record containing all extracted values.</returns>
    private static CopySheetFormatParameters ExtractCopySheetFormatParameters(OperationParameters parameters)
    {
        return new CopySheetFormatParameters(
            parameters.GetOptional("sourceSheetIndex", 0),
            parameters.GetOptional("targetSheetIndex", 0),
            parameters.GetOptional("copyColumnWidths", true),
            parameters.GetOptional("copyRowHeights", true)
        );
    }

    /// <summary>
    ///     Record containing parameters for copy sheet format operations.
    /// </summary>
    /// <param name="SourceSheetIndex">The index of the source sheet.</param>
    /// <param name="TargetSheetIndex">The index of the target sheet.</param>
    /// <param name="CopyColumnWidths">Whether to copy column widths.</param>
    /// <param name="CopyRowHeights">Whether to copy row heights.</param>
    private sealed record CopySheetFormatParameters(
        int SourceSheetIndex,
        int TargetSheetIndex,
        bool CopyColumnWidths,
        bool CopyRowHeights);
}
