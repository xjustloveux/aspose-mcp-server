using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Filter;

/// <summary>
///     Handler for applying filter criteria to Excel worksheets.
/// </summary>
public class FilterByValueExcelHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "filter";

    /// <summary>
    ///     Applies filter criteria to a specific column.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: range, criteria
    ///     Optional: sheetIndex (default: 0), columnIndex (default: 0), filterOperator (default: "Equal")
    /// </param>
    /// <returns>Success message with filter details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var range = parameters.GetRequired<string>("range");
        var criteria = parameters.GetRequired<string>("criteria");
        var columnIndex = parameters.GetOptional("columnIndex", 0);
        var filterOperatorStr = parameters.GetOptional("filterOperator", "Equal");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        ExcelHelper.CreateRange(worksheet.Cells, range);
        worksheet.AutoFilter.Range = range;

        var filterOperator = ExcelFilterHelper.ParseFilterOperator(filterOperatorStr);

        if (filterOperator == FilterOperatorType.Equal)
        {
            worksheet.AutoFilter.Filter(columnIndex, criteria);
        }
        else
        {
            object criteriaValue = criteria;
            if (ExcelFilterHelper.IsNumericOperator(filterOperator) && double.TryParse(criteria, out var numericValue))
                criteriaValue = numericValue;

            worksheet.AutoFilter.Custom(columnIndex, filterOperator, criteriaValue);
        }

        worksheet.AutoFilter.Refresh();

        MarkModified(context);

        return Success(
            $"Filter applied to column {columnIndex} with criteria '{criteria}' (operator: {filterOperatorStr}).");
    }
}
