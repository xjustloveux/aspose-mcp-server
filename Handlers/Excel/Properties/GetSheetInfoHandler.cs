using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Properties;

/// <summary>
///     Handler for getting sheet information from Excel files.
/// </summary>
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
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var getParams = ExtractGetSheetInfoParameters(parameters);

        var workbook = context.Document;

        List<object> sheetList = [];

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

        var result = new
        {
            count = sheetList.Count,
            totalWorksheets = workbook.Worksheets.Count,
            items = sheetList
        };

        return JsonResult(result);
    }

    /// <summary>
    ///     Creates a sheet information object containing worksheet details.
    /// </summary>
    /// <param name="worksheet">The worksheet to get information from.</param>
    /// <param name="index">The index of the worksheet.</param>
    /// <returns>An object containing the worksheet's information.</returns>
    private static object CreateSheetInfo(Worksheet worksheet, int index)
    {
        return new
        {
            index,
            name = worksheet.Name,
            visibility = worksheet.VisibilityType.ToString(),
            dataRowCount = worksheet.Cells.MaxDataRow + 1,
            dataColumnCount = worksheet.Cells.MaxDataColumn + 1,
            usedRange = new
            {
                rowCount = worksheet.Cells.MaxRow + 1,
                columnCount = worksheet.Cells.MaxColumn + 1
            },
            pageOrientation = worksheet.PageSetup.Orientation.ToString(),
            paperSize = worksheet.PageSetup.PaperSize.ToString(),
            freezePanes = new
            {
                row = worksheet.FirstVisibleRow,
                column = worksheet.FirstVisibleColumn
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
    private record GetSheetInfoParameters(int? TargetSheetIndex);
}
