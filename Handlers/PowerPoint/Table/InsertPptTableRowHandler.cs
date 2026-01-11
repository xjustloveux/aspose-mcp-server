using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Table;

/// <summary>
///     Handler for inserting rows into PowerPoint tables.
/// </summary>
public class InsertPptTableRowHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "insert_row";

    /// <summary>
    ///     Inserts a new row into a table.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: shapeIndex.
    ///     Optional: slideIndex (default: 0), rowIndex (default: end), copyFromRow.
    /// </param>
    /// <returns>Success message with insertion details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetOptional("slideIndex", 0);
        var shapeIndex = parameters.GetRequired<int>("shapeIndex");
        var rowIndex = parameters.GetOptional<int?>("rowIndex");
        var copyFromRow = parameters.GetOptional<int?>("copyFromRow");

        var presentation = context.Document;
        var slide = PptTableHelper.GetSlide(presentation, slideIndex);
        var table = PptTableHelper.GetTable(slide, shapeIndex);

        var insertIndex = rowIndex ?? table.Rows.Count;
        if (insertIndex < 0 || insertIndex > table.Rows.Count)
            throw new ArgumentException(
                $"rowIndex must be between 0 and {table.Rows.Count}, got: {insertIndex}");

        var sourceRowIndex = copyFromRow ?? (table.Rows.Count > 0 ? 0 : -1);
        if (sourceRowIndex >= 0 && sourceRowIndex < table.Rows.Count)
            table.Rows.InsertClone(insertIndex, table.Rows[sourceRowIndex], false);
        else
            table.Rows.AddClone(table.Rows[0], false);

        MarkModified(context);

        return Success($"Row inserted at index {insertIndex}.");
    }
}
