using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Table;

/// <summary>
///     Handler for deleting columns from PowerPoint tables.
/// </summary>
public class DeletePptTableColumnHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "delete_column";

    /// <summary>
    ///     Deletes a column from a table.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: shapeIndex, columnIndex
    ///     Optional: slideIndex
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetOptional("slideIndex", 0);
        var shapeIndex = parameters.GetOptional<int?>("shapeIndex");
        var columnIndex = parameters.GetOptional<int?>("columnIndex");

        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for delete_column operation");
        if (!columnIndex.HasValue)
            throw new ArgumentException("columnIndex is required for delete_column operation");

        var presentation = context.Document;
        var slide = PptTableHelper.GetSlide(presentation, slideIndex);
        var table = PptTableHelper.GetTable(slide, shapeIndex.Value);

        table.Columns.RemoveAt(columnIndex.Value, false);

        MarkModified(context);

        return Success($"Column {columnIndex.Value} deleted.");
    }
}
