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
        var setParams = ExtractSetMessagesParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, setParams.SheetIndex);
        var validations = worksheet.Validations;

        ExcelDataValidationHelper.ValidateCollectionIndex(setParams.ValidationIndex, validations.Count,
            "data validation");

        var validation = validations[setParams.ValidationIndex];
        List<string> changes = [];

        if (setParams.ErrorMessage != null)
        {
            validation.ErrorMessage = setParams.ErrorMessage;
            validation.ShowError = !string.IsNullOrEmpty(setParams.ErrorMessage);
            changes.Add($"ErrorMessage={setParams.ErrorMessage}");
        }

        if (setParams.InputMessage != null)
        {
            validation.InputMessage = setParams.InputMessage;
            validation.ShowInput = !string.IsNullOrEmpty(setParams.InputMessage);
            changes.Add($"InputMessage={setParams.InputMessage}");
        }

        MarkModified(context);

        var changesStr = changes.Count > 0 ? string.Join(", ", changes) : "No changes";
        return Success($"Updated data validation #{setParams.ValidationIndex} messages ({changesStr}).");
    }

    private static SetMessagesParameters ExtractSetMessagesParameters(OperationParameters parameters)
    {
        return new SetMessagesParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetRequired<int>("validationIndex"),
            parameters.GetOptional<string?>("errorMessage"),
            parameters.GetOptional<string?>("inputMessage"));
    }

    private record SetMessagesParameters(
        int SheetIndex,
        int ValidationIndex,
        string? ErrorMessage,
        string? InputMessage);
}
