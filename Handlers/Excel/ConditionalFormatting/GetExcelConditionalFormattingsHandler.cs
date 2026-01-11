using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.ConditionalFormatting;

/// <summary>
///     Handler for getting conditional formatting from Excel worksheets.
/// </summary>
public class GetExcelConditionalFormattingsHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all conditional formatting from a worksheet.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>JSON result with conditional formatting information.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var conditionalFormattings = worksheet.ConditionalFormattings;

        if (conditionalFormattings.Count == 0)
            return JsonResult(new
            {
                count = 0,
                sheetIndex,
                worksheetName = worksheet.Name,
                items = Array.Empty<object>(),
                message = "No conditional formattings found"
            });

        List<object> formattingList = [];
        for (var i = 0; i < conditionalFormattings.Count; i++)
        {
            var fcs = conditionalFormattings[i];

            List<string> areasList = [];
            for (var k = 0; k < fcs.RangeCount; k++)
            {
                var area = fcs.GetCellArea(k);
                areasList.Add(
                    $"{CellsHelper.CellIndexToName(area.StartRow, area.StartColumn)}:{CellsHelper.CellIndexToName(area.EndRow, area.EndColumn)}");
            }

            List<object> conditionsList = [];
            for (var j = 0; j < fcs.Count; j++)
            {
                var fc = fcs[j];
                conditionsList.Add(new
                {
                    index = j,
                    operatorType = fc.Operator.ToString(),
                    formula1 = fc.Formula1,
                    formula2 = fc.Formula2,
                    foregroundColor = fc.Style?.ForegroundColor.ToString(),
                    backgroundColor = fc.Style?.BackgroundColor.ToString()
                });
            }

            formattingList.Add(new
            {
                index = i,
                areas = areasList,
                conditionsCount = fcs.Count,
                conditions = conditionsList
            });
        }

        return JsonResult(new
        {
            count = conditionalFormattings.Count,
            sheetIndex,
            worksheetName = worksheet.Name,
            items = formattingList
        });
    }
}
