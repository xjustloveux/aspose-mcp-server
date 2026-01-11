using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.NamedRange;

/// <summary>
///     Handler for deleting named ranges from Excel workbooks.
/// </summary>
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
    public override string Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var name = parameters.GetRequired<string>("name");

        var workbook = context.Document;
        var names = workbook.Worksheets.Names;

        if (names[name] == null)
            throw new ArgumentException($"Named range '{name}' does not exist.");

        names.Remove(name);

        MarkModified(context);

        return Success($"Named range '{name}' deleted.");
    }
}
