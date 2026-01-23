using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.DataValidation;

namespace AsposeMcpServer.Handlers.Excel.DataValidation;

/// <summary>
///     Handler for getting data validation from Excel worksheets.
/// </summary>
[ResultType(typeof(GetDataValidationsResult))]
public class GetExcelDataValidationsHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all data validation information for the worksheet.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>JSON result with data validation information.</returns>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var getParams = ExtractGetParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, getParams.SheetIndex);
        var validations = worksheet.Validations;

        List<DataValidationInfo> validationList = [];
        for (var i = 0; i < validations.Count; i++)
        {
            var validation = validations[i];
            validationList.Add(new DataValidationInfo
            {
                Index = i,
                Type = validation.Type.ToString(),
                OperatorType = validation.Operator.ToString(),
                Formula1 = validation.Formula1,
                Formula2 = validation.Formula2,
                ErrorMessage = validation.ErrorMessage,
                InputMessage = validation.InputMessage,
                ShowError = validation.ShowError,
                ShowInput = validation.ShowInput,
                InCellDropDown = validation.InCellDropDown
            });
        }

        return new GetDataValidationsResult
        {
            Count = validations.Count,
            WorksheetName = worksheet.Name,
            Items = validationList
        };
    }

    private static GetParameters ExtractGetParameters(OperationParameters parameters)
    {
        return new GetParameters(parameters.GetOptional("sheetIndex", 0));
    }

    private sealed record GetParameters(int SheetIndex);
}
