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
        var addParams = ExtractAddParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, addParams.SheetIndex);
        var cells = worksheet.Cells;

        var cellRange = ExcelHelper.CreateRange(cells, addParams.Range);

        var area = new CellArea
        {
            StartRow = cellRange.FirstRow,
            StartColumn = cellRange.FirstColumn,
            EndRow = cellRange.FirstRow + cellRange.RowCount - 1,
            EndColumn = cellRange.FirstColumn + cellRange.ColumnCount - 1
        };
        var validationIndex = worksheet.Validations.Add(area);
        var validation = worksheet.Validations[validationIndex];

        var vType = ExcelDataValidationHelper.ParseValidationType(addParams.ValidationType);
        validation.Type = vType;
        validation.Formula1 = addParams.Formula1;

        if (!string.IsNullOrEmpty(addParams.Formula2))
            validation.Formula2 = addParams.Formula2;

        validation.Operator = ExcelDataValidationHelper.ParseOperatorType(addParams.OperatorType, addParams.Formula2);

        if (!string.IsNullOrEmpty(addParams.ErrorMessage))
        {
            validation.ErrorMessage = addParams.ErrorMessage;
            validation.ShowError = true;
        }

        if (!string.IsNullOrEmpty(addParams.InputMessage))
        {
            validation.InputMessage = addParams.InputMessage;
            validation.ShowInput = true;
        }

        if (vType == ValidationType.List)
            validation.InCellDropDown = addParams.InCellDropDown;

        MarkModified(context);

        return Success(
            $"Data validation added to range {addParams.Range} (type: {addParams.ValidationType}, index: {validationIndex}).");
    }

    private static AddParameters ExtractAddParameters(OperationParameters parameters)
    {
        return new AddParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetRequired<string>("range"),
            parameters.GetRequired<string>("validationType"),
            parameters.GetRequired<string>("formula1"),
            parameters.GetOptional<string?>("formula2"),
            parameters.GetOptional<string?>("operatorType"),
            parameters.GetOptional("inCellDropDown", true),
            parameters.GetOptional<string?>("errorMessage"),
            parameters.GetOptional<string?>("inputMessage"));
    }

    private sealed record AddParameters(
        int SheetIndex,
        string Range,
        string ValidationType,
        string Formula1,
        string? Formula2,
        string? OperatorType,
        bool InCellDropDown,
        string? ErrorMessage,
        string? InputMessage);
}
