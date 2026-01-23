using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.NamedRange;

/// <summary>
///     Handler for adding named ranges to Excel workbooks.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractAddNamedRangeParameters(parameters);

        var workbook = context.Document;
        var names = workbook.Worksheets.Names;

        if (names[p.Name] != null)
            throw new ArgumentException($"Named range '{p.Name}' already exists.");

        try
        {
            Aspose.Cells.Range rangeObject;

            if (p.Range.Contains('!'))
            {
                rangeObject = ExcelNamedRangeHelper.ParseRangeWithSheetReference(workbook, p.Range);
            }
            else
            {
                var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);
                rangeObject = ExcelNamedRangeHelper.CreateRangeFromAddress(worksheet.Cells, p.Range);
            }

            rangeObject.Name = p.Name;

            var namedRange = names[p.Name];
            if (!string.IsNullOrEmpty(p.Comment))
                namedRange.Comment = p.Comment;

            MarkModified(context);

            return new SuccessResult { Message = $"Named range '{p.Name}' added (reference: {namedRange.RefersTo})." };
        }
        catch (ArgumentException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                $"Failed to create named range '{p.Name}' with range '{p.Range}': {ex.Message}", ex);
        }
    }

    private static AddNamedRangeParameters ExtractAddNamedRangeParameters(OperationParameters parameters)
    {
        return new AddNamedRangeParameters(
            parameters.GetOptional("sheetIndex", 0),
            parameters.GetRequired<string>("name"),
            parameters.GetRequired<string>("range"),
            parameters.GetOptional<string?>("comment")
        );
    }

    private sealed record AddNamedRangeParameters(int SheetIndex, string Name, string Range, string? Comment);
}
