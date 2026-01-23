using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Filter;

/// <summary>
///     Handler for applying filter criteria to Excel worksheets.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var filterParams = ExtractFilterByValueParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, filterParams.SheetIndex);

        ExcelHelper.CreateRange(worksheet.Cells, filterParams.Range);
        worksheet.AutoFilter.Range = filterParams.Range;

        var filterOperator = ExcelFilterHelper.ParseFilterOperator(filterParams.FilterOperator);

        if (filterOperator == FilterOperatorType.Equal)
        {
            worksheet.AutoFilter.Filter(filterParams.ColumnIndex, filterParams.Criteria);
        }
        else
        {
            object criteriaValue = filterParams.Criteria;
            if (ExcelFilterHelper.IsNumericOperator(filterOperator) &&
                double.TryParse(filterParams.Criteria, out var numericValue))
                criteriaValue = numericValue;

            worksheet.AutoFilter.Custom(filterParams.ColumnIndex, filterOperator, criteriaValue);
        }

        worksheet.AutoFilter.Refresh();

        MarkModified(context);

        return new SuccessResult
        {
            Message =
                $"Filter applied to column {filterParams.ColumnIndex} with criteria '{filterParams.Criteria}' (operator: {filterParams.FilterOperator})."
        };
    }

    /// <summary>
    ///     Extracts filter by value parameters from the operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted filter by value parameters.</returns>
    private static FilterByValueParameters ExtractFilterByValueParameters(OperationParameters parameters)
    {
        return new FilterByValueParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetRequired<string>("range"),
            parameters.GetRequired<string>("criteria"),
            parameters.GetOptional("columnIndex", 0),
            parameters.GetOptional("filterOperator", "Equal")
        );
    }

    /// <summary>
    ///     Parameters for filter by value operation.
    /// </summary>
    /// <param name="SheetIndex">The worksheet index (0-based).</param>
    /// <param name="Range">The range to apply filter to.</param>
    /// <param name="Criteria">The filter criteria value.</param>
    /// <param name="ColumnIndex">The column index to filter (0-based).</param>
    /// <param name="FilterOperator">The filter operator type.</param>
    private sealed record FilterByValueParameters(
        int SheetIndex,
        string Range,
        string Criteria,
        int ColumnIndex,
        string FilterOperator);
}
