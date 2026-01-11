using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Table;

/// <summary>
///     Handler for deleting rows from PowerPoint tables.
/// </summary>
public class DeletePptTableRowHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "delete_row";

    /// <summary>
    ///     Deletes a row from a table.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: shapeIndex, rowIndex
    ///     Optional: slideIndex
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetOptional("slideIndex", 0);
        var shapeIndex = parameters.GetOptional<int?>("shapeIndex");
        var rowIndex = parameters.GetOptional<int?>("rowIndex");

        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for delete_row operation");
        if (!rowIndex.HasValue)
            throw new ArgumentException("rowIndex is required for delete_row operation");

        var presentation = context.Document;
        var slide = PptTableHelper.GetSlide(presentation, slideIndex);
        var table = PptTableHelper.GetTable(slide, shapeIndex.Value);

        table.Rows.RemoveAt(rowIndex.Value, false);

        MarkModified(context);

        return Success($"Row {rowIndex.Value} deleted.");
    }
}
