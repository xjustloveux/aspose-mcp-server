using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.DataValidation;

/// <summary>
///     Handler for adding data validation to Excel worksheets.
/// </summary>
public class AddExcelDataValidationHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds data validation to a range.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: range, validationType, formula1
    ///     Optional: sheetIndex (default: 0), formula2, operatorType, inCellDropDown, errorMessage, inputMessage
    /// </param>
    /// <returns>Success message with validation details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var range = parameters.GetRequired<string>("range");
        var validationType = parameters.GetRequired<string>("validationType");
        var formula1 = parameters.GetRequired<string>("formula1");
        var formula2 = parameters.GetOptional<string?>("formula2");
        var operatorType = parameters.GetOptional<string?>("operatorType");
        var inCellDropDown = parameters.GetOptional("inCellDropDown", true);
        var errorMessage = parameters.GetOptional<string?>("errorMessage");
        var inputMessage = parameters.GetOptional<string?>("inputMessage");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var cells = worksheet.Cells;

        var cellRange = ExcelHelper.CreateRange(cells, range);

        var area = new CellArea
        {
            StartRow = cellRange.FirstRow,
            StartColumn = cellRange.FirstColumn,
            EndRow = cellRange.FirstRow + cellRange.RowCount - 1,
            EndColumn = cellRange.FirstColumn + cellRange.ColumnCount - 1
        };
        var validationIndex = worksheet.Validations.Add(area);
        var validation = worksheet.Validations[validationIndex];

        var vType = ExcelDataValidationHelper.ParseValidationType(validationType);
        validation.Type = vType;
        validation.Formula1 = formula1;

        if (!string.IsNullOrEmpty(formula2))
            validation.Formula2 = formula2;

        validation.Operator = ExcelDataValidationHelper.ParseOperatorType(operatorType, formula2);

        if (!string.IsNullOrEmpty(errorMessage))
        {
            validation.ErrorMessage = errorMessage;
            validation.ShowError = true;
        }

        if (!string.IsNullOrEmpty(inputMessage))
        {
            validation.InputMessage = inputMessage;
            validation.ShowInput = true;
        }

        if (vType == ValidationType.List)
            validation.InCellDropDown = inCellDropDown;

        MarkModified(context);

        return Success($"Data validation added to range {range} (type: {validationType}, index: {validationIndex}).");
    }
}
