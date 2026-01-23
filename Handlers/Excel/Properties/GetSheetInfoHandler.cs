using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.Properties;

namespace AsposeMcpServer.Handlers.Excel.Properties;

/// <summary>
///     Handler for getting sheet information from Excel files.
/// </summary>
[ResultType(typeof(GetSheetInfoResult))]
public class GetSheetInfoHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get_sheet_info";

    /// <summary>
    ///     Gets information about all worksheets or a specific worksheet.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: targetSheetIndex (if not provided, returns all sheets)
    /// </param>
    /// <returns>JSON result with sheet information.</returns>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var getParams = ExtractGetSheetInfoParameters(parameters);

        var workbook = context.Document;

        List<SheetInfoDetail> sheetList = [];

        if (getParams.TargetSheetIndex.HasValue)
        {
            var worksheet = ExcelHelper.GetWorksheet(workbook, getParams.TargetSheetIndex.Value);
            sheetList.Add(CreateSheetInfo(worksheet, getParams.TargetSheetIndex.Value));
        }
        else
        {
            for (var i = 0; i < workbook.Worksheets.Count; i++)
                sheetList.Add(CreateSheetInfo(workbook.Worksheets[i], i));
        }

        return new GetSheetInfoResult
        {
            Count = sheetList.Count,
            TotalWorksheets = workbook.Worksheets.Count,
            Items = sheetList
        };
    }

    /// <summary>
    ///     Creates a sheet information object containing worksheet details.
    /// </summary>
    /// <param name="worksheet">The worksheet to get information from.</param>
    /// <param name="index">The index of the worksheet.</param>
    /// <returns>A SheetInfoDetail object containing the worksheet's information.</returns>
    private static SheetInfoDetail CreateSheetInfo(Worksheet worksheet, int index)
    {
        return new SheetInfoDetail
        {
            Index = index,
            Name = worksheet.Name,
            Visibility = worksheet.VisibilityType.ToString(),
            DataRowCount = worksheet.Cells.MaxDataRow + 1,
            DataColumnCount = worksheet.Cells.MaxDataColumn + 1,
            UsedRange = new UsedRangeInfo
            {
                RowCount = worksheet.Cells.MaxRow + 1,
                ColumnCount = worksheet.Cells.MaxColumn + 1
            },
            PageOrientation = worksheet.PageSetup.Orientation.ToString(),
            PaperSize = worksheet.PageSetup.PaperSize.ToString(),
            FreezePanes = new FreezePanesInfo
            {
                Row = worksheet.FirstVisibleRow,
                Column = worksheet.FirstVisibleColumn
            }
        };
    }

    /// <summary>
    ///     Extracts get sheet info parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted get sheet info parameters.</returns>
    private static GetSheetInfoParameters ExtractGetSheetInfoParameters(OperationParameters parameters)
    {
        return new GetSheetInfoParameters(
            parameters.GetOptional<int?>("targetSheetIndex")
        );
    }

    /// <summary>
    ///     Record to hold get sheet info parameters.
    /// </summary>
    private sealed record GetSheetInfoParameters(int? TargetSheetIndex);
}
