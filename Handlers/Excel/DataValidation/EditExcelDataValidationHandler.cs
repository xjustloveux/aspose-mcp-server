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
        var editParams = ExtractEditParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, editParams.SheetIndex);
        var validations = worksheet.Validations;

        ExcelDataValidationHelper.ValidateCollectionIndex(editParams.ValidationIndex, validations.Count,
            "data validation");

        var validation = validations[editParams.ValidationIndex];
        List<string> changes = [];

        if (!string.IsNullOrEmpty(editParams.ValidationType))
        {
            validation.Type = ExcelDataValidationHelper.ParseValidationType(editParams.ValidationType);
            changes.Add($"Type={editParams.ValidationType}");
        }

        if (!string.IsNullOrEmpty(editParams.Formula1))
        {
            validation.Formula1 = editParams.Formula1;
            changes.Add($"Formula1={editParams.Formula1}");
        }

        if (editParams.Formula2 != null)
        {
            validation.Formula2 = editParams.Formula2;
            changes.Add($"Formula2={editParams.Formula2}");
        }

        if (!string.IsNullOrEmpty(editParams.OperatorType))
        {
            validation.Operator =
                ExcelDataValidationHelper.ParseOperatorType(editParams.OperatorType, editParams.Formula2);
            changes.Add($"Operator={editParams.OperatorType}");
        }

        if (editParams.InCellDropDown.HasValue)
        {
            validation.InCellDropDown = editParams.InCellDropDown.Value;
            changes.Add($"InCellDropDown={editParams.InCellDropDown.Value}");
        }

        if (editParams.ErrorMessage != null)
        {
            validation.ErrorMessage = editParams.ErrorMessage;
            validation.ShowError = !string.IsNullOrEmpty(editParams.ErrorMessage);
            changes.Add($"ErrorMessage={editParams.ErrorMessage}");
        }

        if (editParams.InputMessage != null)
        {
            validation.InputMessage = editParams.InputMessage;
            validation.ShowInput = !string.IsNullOrEmpty(editParams.InputMessage);
            changes.Add($"InputMessage={editParams.InputMessage}");
        }

        MarkModified(context);

        var changesStr = changes.Count > 0 ? string.Join(", ", changes) : "No changes";
        return Success($"Edited data validation #{editParams.ValidationIndex} ({changesStr}).");
    }

    private static EditParameters ExtractEditParameters(OperationParameters parameters)
    {
        return new EditParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetRequired<int>("validationIndex"),
            parameters.GetOptional<string?>("validationType"),
            parameters.GetOptional<string?>("formula1"),
            parameters.GetOptional<string?>("formula2"),
            parameters.GetOptional<string?>("operatorType"),
            parameters.GetOptional<bool?>("inCellDropDown"),
            parameters.GetOptional<string?>("errorMessage"),
            parameters.GetOptional<string?>("inputMessage"));
    }

    private record EditParameters(
        int SheetIndex,
        int ValidationIndex,
        string? ValidationType,
        string? Formula1,
        string? Formula2,
        string? OperatorType,
        bool? InCellDropDown,
        string? ErrorMessage,
        string? InputMessage);
}
