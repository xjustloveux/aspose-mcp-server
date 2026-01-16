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
    /// <exception cref="ArgumentException">Thrown when incorrect password is provided.</exception>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractUnprotectExcelParameters(parameters);

        var workbook = context.Document;

        if (p.SheetIndex.HasValue)
        {
            var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex.Value);

            if (!worksheet.IsProtected)
            {
                MarkModified(context);
                return Success($"Worksheet '{worksheet.Name}' is not protected.");
            }

            try
            {
                worksheet.Unprotect(p.Password);
            }
            catch (Exception ex)
            {
                throw new ArgumentException(
                    $"Incorrect password. Cannot unprotect worksheet '{worksheet.Name}'. Error: {ex.Message}");
            }

            MarkModified(context);
            return Success($"Worksheet '{worksheet.Name}' protection removed successfully.");
        }

        workbook.Unprotect(p.Password);
        MarkModified(context);
        return Success("Workbook protection removed successfully.");
    }

    /// <summary>
    ///     Extracts parameters for UnprotectExcel operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>Extracted parameters.</returns>
    private static UnprotectExcelParameters ExtractUnprotectExcelParameters(OperationParameters parameters)
    {
        return new UnprotectExcelParameters(
            parameters.GetOptional<int?>("sheetIndex"),
            parameters.GetOptional<string?>("password")
        );
    }

    /// <summary>
    ///     Parameters for UnprotectExcel operation.
    /// </summary>
    /// <param name="SheetIndex">The sheet index (optional).</param>
    /// <param name="Password">The password for unprotection (optional).</param>
    private sealed record UnprotectExcelParameters(int? SheetIndex, string? Password);
}
