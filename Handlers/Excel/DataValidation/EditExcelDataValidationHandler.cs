using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.DataValidation;

/// <summary>
///     Handler for editing data validation in Excel worksheets.
/// </summary>
public class EditExcelDataValidationHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits existing data validation.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: validationIndex
    ///     Optional: sheetIndex (default: 0), validationType, formula1, formula2, operatorType, inCellDropDown, errorMessage,
    ///     inputMessage
    /// </param>
    /// <returns>Success message with changes details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var validationIndex = parameters.GetRequired<int>("validationIndex");
        var validationType = parameters.GetOptional<string?>("validationType");
        var formula1 = parameters.GetOptional<string?>("formula1");
        var formula2 = parameters.GetOptional<string?>("formula2");
        var operatorType = parameters.GetOptional<string?>("operatorType");
        var inCellDropDown = parameters.GetOptional<bool?>("inCellDropDown");
        var errorMessage = parameters.GetOptional<string?>("errorMessage");
        var inputMessage = parameters.GetOptional<string?>("inputMessage");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var validations = worksheet.Validations;

        ExcelDataValidationHelper.ValidateCollectionIndex(validationIndex, validations.Count, "data validation");

        var validation = validations[validationIndex];
        List<string> changes = [];

        if (!string.IsNullOrEmpty(validationType))
        {
            validation.Type = ExcelDataValidationHelper.ParseValidationType(validationType);
            changes.Add($"Type={validationType}");
        }

        if (!string.IsNullOrEmpty(formula1))
        {
            validation.Formula1 = formula1;
            changes.Add($"Formula1={formula1}");
        }

        if (formula2 != null)
        {
            validation.Formula2 = formula2;
            changes.Add($"Formula2={formula2}");
        }

        if (!string.IsNullOrEmpty(operatorType))
        {
            validation.Operator = ExcelDataValidationHelper.ParseOperatorType(operatorType, formula2);
            changes.Add($"Operator={operatorType}");
        }

        if (inCellDropDown.HasValue)
        {
            validation.InCellDropDown = inCellDropDown.Value;
            changes.Add($"InCellDropDown={inCellDropDown.Value}");
        }

        if (errorMessage != null)
        {
            validation.ErrorMessage = errorMessage;
            validation.ShowError = !string.IsNullOrEmpty(errorMessage);
            changes.Add($"ErrorMessage={errorMessage}");
        }

        if (inputMessage != null)
        {
            validation.InputMessage = inputMessage;
            validation.ShowInput = !string.IsNullOrEmpty(inputMessage);
            changes.Add($"InputMessage={inputMessage}");
        }

        MarkModified(context);

        var changesStr = changes.Count > 0 ? string.Join(", ", changes) : "No changes";
        return Success($"Edited data validation #{validationIndex} ({changesStr}).");
    }
}
