using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Table;

/// <summary>
///     Handler for inserting columns into PowerPoint tables.
/// </summary>
public class InsertPptTableColumnHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "insert_column";

    /// <summary>
    ///     Inserts a new column into a table.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: shapeIndex.
    ///     Optional: slideIndex (default: 0), columnIndex (default: end), copyFromColumn.
    /// </param>
    /// <returns>Success message with insertion details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetOptional("slideIndex", 0);
        var shapeIndex = parameters.GetRequired<int>("shapeIndex");
        var columnIndex = parameters.GetOptional<int?>("columnIndex");
        var copyFromColumn = parameters.GetOptional<int?>("copyFromColumn");

        var presentation = context.Document;
        var slide = PptTableHelper.GetSlide(presentation, slideIndex);
        var table = PptTableHelper.GetTable(slide, shapeIndex);

        var insertIndex = columnIndex ?? table.Columns.Count;
        if (insertIndex < 0 || insertIndex > table.Columns.Count)
            throw new ArgumentException(
                $"columnIndex must be between 0 and {table.Columns.Count}, got: {insertIndex}");

        var sourceColIndex = copyFromColumn ?? (table.Columns.Count > 0 ? 0 : -1);
        if (sourceColIndex >= 0 && sourceColIndex < table.Columns.Count)
            table.Columns.InsertClone(insertIndex, table.Columns[sourceColIndex], false);
        else
            table.Columns.AddClone(table.Columns[0], false);

        MarkModified(context);

        return Success($"Column inserted at index {insertIndex}.");
    }
}
