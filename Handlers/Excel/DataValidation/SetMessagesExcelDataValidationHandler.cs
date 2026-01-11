using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.DataValidation;

/// <summary>
///     Handler for setting input/error messages on data validation in Excel worksheets.
/// </summary>
public class SetMessagesExcelDataValidationHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "set_messages";

    /// <summary>
    ///     Sets input/error messages for data validation.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: validationIndex
    ///     Optional: sheetIndex (default: 0), errorMessage, inputMessage
    /// </param>
    /// <returns>Success message with changes details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var validationIndex = parameters.GetRequired<int>("validationIndex");
        var errorMessage = parameters.GetOptional<string?>("errorMessage");
        var inputMessage = parameters.GetOptional<string?>("inputMessage");

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
        var validations = worksheet.Validations;

        ExcelDataValidationHelper.ValidateCollectionIndex(validationIndex, validations.Count, "data validation");

        var validation = validations[validationIndex];
        List<string> changes = [];

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
        return Success($"Updated data validation #{validationIndex} messages ({changesStr}).");
    }
}
