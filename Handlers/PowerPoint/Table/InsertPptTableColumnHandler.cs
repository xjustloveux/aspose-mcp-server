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
        var insertParams = ExtractInsertColumnParameters(parameters);

        var presentation = context.Document;
        var slide = PptTableHelper.GetSlide(presentation, insertParams.SlideIndex);
        var table = PptTableHelper.GetTable(slide, insertParams.ShapeIndex);

        var insertIndex = insertParams.ColumnIndex ?? table.Columns.Count;
        if (insertIndex < 0 || insertIndex > table.Columns.Count)
            throw new ArgumentException(
                $"columnIndex must be between 0 and {table.Columns.Count}, got: {insertIndex}");

        var sourceColIndex = insertParams.CopyFromColumn ?? (table.Columns.Count > 0 ? 0 : -1);
        if (sourceColIndex >= 0 && sourceColIndex < table.Columns.Count)
            table.Columns.InsertClone(insertIndex, table.Columns[sourceColIndex], false);
        else
            table.Columns.AddClone(table.Columns[0], false);

        MarkModified(context);

        return Success($"Column inserted at index {insertIndex}.");
    }

    /// <summary>
    ///     Extracts insert column parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted insert column parameters.</returns>
    private static InsertColumnParameters ExtractInsertColumnParameters(OperationParameters parameters)
    {
        return new InsertColumnParameters(
            parameters.GetOptional("slideIndex", 0),
            parameters.GetRequired<int>("shapeIndex"),
            parameters.GetOptional<int?>("columnIndex"),
            parameters.GetOptional<int?>("copyFromColumn")
        );
    }

    /// <summary>
    ///     Record for holding insert column parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ShapeIndex">The shape index.</param>
    /// <param name="ColumnIndex">The optional column index to insert at.</param>
    /// <param name="CopyFromColumn">The optional column to copy from.</param>
    private sealed record InsertColumnParameters(int SlideIndex, int ShapeIndex, int? ColumnIndex, int? CopyFromColumn);
}
