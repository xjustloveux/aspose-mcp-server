using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.NamedRange;

/// <summary>
///     Handler for deleting named ranges from Excel workbooks.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class DeleteExcelNamedRangeHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a named range from the workbook.
    /// </summary>
    /// <param name="context">The workbook context.</param>
    /// <param name="parameters">
    ///     Required: name
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractDeleteNamedRangeParameters(parameters);

        var workbook = context.Document;
        var names = workbook.Worksheets.Names;

        if (names[p.Name] == null)
            throw new ArgumentException($"Named range '{p.Name}' does not exist.");

        names.Remove(p.Name);

        MarkModified(context);

        return new SuccessResult { Message = $"Named range '{p.Name}' deleted." };
    }

    private static DeleteNamedRangeParameters ExtractDeleteNamedRangeParameters(OperationParameters parameters)
    {
        return new DeleteNamedRangeParameters(
            parameters.GetRequired<string>("name")
        );
    }

    private sealed record DeleteNamedRangeParameters(string Name);
}
