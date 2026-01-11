using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.Excel.NamedRange;

/// <summary>
///     Handler for adding named ranges to Excel workbooks.
/// </summary>
public class AddExcelNamedRangeHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a named range to the workbook.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: name, range
    ///     Optional: sheetIndex (default: 0), comment
    /// </param>
    /// <returns>Success message with named range details.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var name = parameters.GetRequired<string>("name");
        var range = parameters.GetRequired<string>("range");
        var comment = parameters.GetOptional<string?>("comment");

        var workbook = context.Document;
        var names = workbook.Worksheets.Names;

        if (names[name] != null)
            throw new ArgumentException($"Named range '{name}' already exists.");

        try
        {
            Aspose.Cells.Range rangeObject;

            if (range.Contains('!'))
            {
                rangeObject = ExcelNamedRangeHelper.ParseRangeWithSheetReference(workbook, range);
            }
            else
            {
                var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
                rangeObject = ExcelNamedRangeHelper.CreateRangeFromAddress(worksheet.Cells, range);
            }

            rangeObject.Name = name;

            var namedRange = names[name];
            if (!string.IsNullOrEmpty(comment))
                namedRange.Comment = comment;

            MarkModified(context);

            return Success($"Named range '{name}' added (reference: {namedRange.RefersTo}).");
        }
        catch (ArgumentException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                $"Failed to create named range '{name}' with range '{range}': {ex.Message}", ex);
        }
    }
}
