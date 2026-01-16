using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.DataValidation;

/// <summary>
///     Handler for deleting data validation from Excel worksheets.
/// </summary>
public class DeleteExcelDataValidationHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes data validation by index.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: validationIndex
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var deleteParams = ExtractDeleteParameters(parameters);

        var workbook = context.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, deleteParams.SheetIndex);
        var validations = worksheet.Validations;

        ExcelDataValidationHelper.ValidateCollectionIndex(deleteParams.ValidationIndex, validations.Count,
            "data validation");

        validations.RemoveAt(deleteParams.ValidationIndex);

        MarkModified(context);

        return Success($"Deleted data validation #{deleteParams.ValidationIndex} (remaining: {validations.Count}).");
    }

    private static DeleteParameters ExtractDeleteParameters(OperationParameters parameters)
    {
        return new DeleteParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetRequired<int>("validationIndex"));
    }

    private record DeleteParameters(int SheetIndex, int ValidationIndex);
}
