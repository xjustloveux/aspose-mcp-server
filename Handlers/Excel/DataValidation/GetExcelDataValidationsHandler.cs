using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.DataValidation;

/// <summary>
///     Handler for getting data validation from Excel worksheets.
/// </summary>
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
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var getParams = ExtractGetParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, getParams.SheetIndex);
        var validations = worksheet.Validations;

        List<object> validationList = [];
        for (var i = 0; i < validations.Count; i++)
        {
            var validation = validations[i];
            validationList.Add(new
            {
                index = i,
                type = validation.Type.ToString(),
                operatorType = validation.Operator.ToString(),
                formula1 = validation.Formula1,
                formula2 = validation.Formula2,
                errorMessage = validation.ErrorMessage,
                inputMessage = validation.InputMessage,
                showError = validation.ShowError,
                showInput = validation.ShowInput,
                inCellDropDown = validation.InCellDropDown
            });
        }

        return JsonResult(new
        {
            count = validations.Count,
            worksheetName = worksheet.Name,
            items = validationList
        });
    }

    private static GetParameters ExtractGetParameters(OperationParameters parameters)
    {
        return new GetParameters(parameters.GetOptional("sheetIndex", 0));
    }

    private sealed record GetParameters(int SheetIndex);
}
