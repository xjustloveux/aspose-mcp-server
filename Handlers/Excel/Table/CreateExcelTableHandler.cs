using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Table;

/// <summary>
///     Handler for creating a table (ListObject) in an Excel worksheet.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class CreateExcelTableHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "create";

    /// <summary>
    ///     Creates a table from a cell range.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: range
    ///     Optional: sheetIndex, hasHeaders (default: true), name
    /// </param>
    /// <returns>Success message with table details.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractParameters(parameters);

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);

            var rangeParts = p.Range.Split(':');
            if (rangeParts.Length != 2)
                throw new ArgumentException(
                    $"Invalid range format '{p.Range}'. Expected format: 'A1:D10' (startCell:endCell).");

            var listObjectIndex = worksheet.ListObjects.Add(rangeParts[0].Trim(), rangeParts[1].Trim(), p.HasHeaders);
            var listObject = worksheet.ListObjects[listObjectIndex];

            if (!string.IsNullOrEmpty(p.Name))
                listObject.DisplayName = p.Name;

            MarkModified(context);

            var name = listObject.DisplayName ?? $"Table{listObjectIndex + 1}";
            return new SuccessResult
            {
                Message =
                    $"Table '{name}' created in sheet {p.SheetIndex} with range {p.Range} (hasHeaders={p.HasHeaders})."
            };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Failed to create table with range '{p.Range}': {ex.Message}");
        }
    }

    private static CreateParameters ExtractParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var range = parameters.GetOptional<string?>("range");
        var hasHeaders = parameters.GetOptional("hasHeaders", true);
        var name = parameters.GetOptional<string?>("name");

        if (string.IsNullOrEmpty(range))
            throw new ArgumentException("range is required for create operation");

        return new CreateParameters(sheetIndex, range, hasHeaders, name);
    }

    private sealed record CreateParameters(int SheetIndex, string Range, bool HasHeaders, string? Name);
}
