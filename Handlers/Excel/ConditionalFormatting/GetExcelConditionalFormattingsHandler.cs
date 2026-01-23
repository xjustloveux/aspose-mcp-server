using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.ConditionalFormatting;

namespace AsposeMcpServer.Handlers.Excel.ConditionalFormatting;

/// <summary>
///     Handler for getting conditional formatting from Excel worksheets.
/// </summary>
[ResultType(typeof(GetConditionalFormattingsResult))]
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
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var getParams = ExtractGetParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, getParams.SheetIndex);
        var conditionalFormattings = worksheet.ConditionalFormattings;

        if (conditionalFormattings.Count == 0)
            return new GetConditionalFormattingsResult
            {
                Count = 0,
                SheetIndex = getParams.SheetIndex,
                WorksheetName = worksheet.Name,
                Items = Array.Empty<ConditionalFormattingInfo>(),
                Message = "No conditional formattings found"
            };

        List<ConditionalFormattingInfo> formattingList = [];
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

            List<ConditionalFormattingCondition> conditionsList = [];
            for (var j = 0; j < fcs.Count; j++)
            {
                var fc = fcs[j];
                conditionsList.Add(new ConditionalFormattingCondition
                {
                    Index = j,
                    OperatorType = fc.Operator.ToString(),
                    Formula1 = fc.Formula1,
                    Formula2 = fc.Formula2,
                    ForegroundColor = fc.Style?.ForegroundColor.ToString(),
                    BackgroundColor = fc.Style?.BackgroundColor.ToString()
                });
            }

            formattingList.Add(new ConditionalFormattingInfo
            {
                Index = i,
                Areas = areasList,
                ConditionsCount = fcs.Count,
                Conditions = conditionsList
            });
        }

        return new GetConditionalFormattingsResult
        {
            Count = conditionalFormattings.Count,
            SheetIndex = getParams.SheetIndex,
            WorksheetName = worksheet.Name,
            Items = formattingList
        };
    }

    private static GetParameters ExtractGetParameters(OperationParameters parameters)
    {
        return new GetParameters(parameters.GetOptional("sheetIndex", 0));
    }

    private sealed record GetParameters(int SheetIndex);
}
