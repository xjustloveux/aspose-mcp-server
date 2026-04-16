using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Ole;
using AsposeMcpServer.Results.Shared.Ole;

namespace AsposeMcpServer.Handlers.Excel.OleObject;

/// <summary>
///     Handler for the <c>list</c> operation on <c>excel_ole_object</c>. Enumerates OLE
///     objects across every worksheet (flat index) and projects to the cross-tool shape.
/// </summary>
[ResultType(typeof(OleListResult))]
public sealed class ListExcelOleObjectHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "list";

    /// <summary>Executes the <c>list</c> operation.</summary>
    /// <param name="context">Operation context; <c>Workbook</c> must be non-null.</param>
    /// <param name="parameters">Unused.</param>
    /// <returns>An <see cref="OleListResult" /> spanning all sheets.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="context" /> is null.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        ArgumentNullException.ThrowIfNull(context);
        var workbook = context.Document;
        var items = new List<OleObjectMetadata>();
        var flatIndex = 0;

        for (var si = 0; si < workbook.Worksheets.Count; si++)
        {
            var sheet = workbook.Worksheets[si];
            foreach (var ole in sheet.OleObjects)
            {
                items.Add(ExcelOleMetadataMapper.Map(ole, sheet, si, flatIndex));
                flatIndex++;
            }
        }

        return new OleListResult { Count = items.Count, Items = items };
    }
}
