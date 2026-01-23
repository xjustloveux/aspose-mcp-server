using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Excel.NamedRange;

namespace AsposeMcpServer.Handlers.Excel.NamedRange;

/// <summary>
///     Handler for getting named ranges from Excel workbooks.
/// </summary>
[ResultType(typeof(GetNamedRangesResult))]
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
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var workbook = context.Document;
        var names = workbook.Worksheets.Names;

        if (names.Count == 0)
            return new GetNamedRangesResult
            {
                Count = 0,
                Items = Array.Empty<NamedRangeInfo>(),
                Message = "No named ranges found"
            };

        List<NamedRangeInfo> nameList = [];
        for (var i = 0; i < names.Count; i++)
        {
            var namedRange = names[i];
            nameList.Add(new NamedRangeInfo
            {
                Index = i,
                Name = namedRange.Text,
                Reference = namedRange.RefersTo,
                Comment = namedRange.Comment,
                IsVisible = namedRange.IsVisible
            });
        }

        return new GetNamedRangesResult
        {
            Count = names.Count,
            Items = nameList
        };
    }
}
