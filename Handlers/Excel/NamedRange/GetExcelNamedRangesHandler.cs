using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.NamedRange;

/// <summary>
///     Handler for getting named ranges from Excel workbooks.
/// </summary>
public class GetExcelNamedRangesHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all named ranges from the workbook.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">No parameters required.</param>
    /// <returns>JSON result with named ranges information.</returns>
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var workbook = context.Document;
        var names = workbook.Worksheets.Names;

        if (names.Count == 0)
            return JsonResult(new
            {
                count = 0,
                items = Array.Empty<object>(),
                message = "No named ranges found"
            });

        List<object> nameList = [];
        for (var i = 0; i < names.Count; i++)
        {
            var namedRange = names[i];
            nameList.Add(new
            {
                index = i,
                name = namedRange.Text,
                reference = namedRange.RefersTo,
                comment = namedRange.Comment,
                isVisible = namedRange.IsVisible
            });
        }

        return JsonResult(new
        {
            count = names.Count,
            items = nameList
        });
    }
}
