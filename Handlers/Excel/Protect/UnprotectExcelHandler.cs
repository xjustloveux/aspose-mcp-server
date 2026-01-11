using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.Protect;

/// <summary>
///     Handler for removing protection from Excel workbook or worksheet.
/// </summary>
public class UnprotectExcelHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "unprotect";

    /// <summary>
    ///     Removes protection from workbook or worksheet.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex, password
    /// </param>
    /// <returns>Success message with unprotection details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional<int?>("sheetIndex");
        var password = parameters.GetOptional<string?>("password");

        var workbook = context.Document;

        if (sheetIndex.HasValue)
        {
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex.Value);

            if (!worksheet.IsProtected)
            {
                MarkModified(context);
                return Success($"Worksheet '{worksheet.Name}' is not protected.");
            }

            try
            {
                worksheet.Unprotect(password);
            }
            catch (Exception ex)
            {
                throw new ArgumentException(
                    $"Incorrect password. Cannot unprotect worksheet '{worksheet.Name}'. Error: {ex.Message}");
            }

            MarkModified(context);
            return Success($"Worksheet '{worksheet.Name}' protection removed successfully.");
        }

        workbook.Unprotect(password);
        MarkModified(context);
        return Success("Workbook protection removed successfully.");
    }
}
