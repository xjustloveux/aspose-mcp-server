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
        var insertParams = ExtractInsertRowParameters(parameters);

        var presentation = context.Document;
        var slide = PptTableHelper.GetSlide(presentation, insertParams.SlideIndex);
        var table = PptTableHelper.GetTable(slide, insertParams.ShapeIndex);

        var insertIndex = insertParams.RowIndex ?? table.Rows.Count;
        if (insertIndex < 0 || insertIndex > table.Rows.Count)
            throw new ArgumentException(
                $"rowIndex must be between 0 and {table.Rows.Count}, got: {insertIndex}");

        var sourceRowIndex = insertParams.CopyFromRow ?? (table.Rows.Count > 0 ? 0 : -1);
        if (sourceRowIndex >= 0 && sourceRowIndex < table.Rows.Count)
            table.Rows.InsertClone(insertIndex, table.Rows[sourceRowIndex], false);
        else
            table.Rows.AddClone(table.Rows[0], false);

        MarkModified(context);

        return Success($"Row inserted at index {insertIndex}.");
    }

    /// <summary>
    ///     Extracts insert row parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted insert row parameters.</returns>
    private static InsertRowParameters ExtractInsertRowParameters(OperationParameters parameters)
    {
        return new InsertRowParameters(
            parameters.GetOptional("slideIndex", 0),
            parameters.GetRequired<int>("shapeIndex"),
            parameters.GetOptional<int?>("rowIndex"),
            parameters.GetOptional<int?>("copyFromRow")
        );
    }

    /// <summary>
    ///     Record for holding insert row parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ShapeIndex">The shape index.</param>
    /// <param name="RowIndex">The optional row index to insert at.</param>
    /// <param name="CopyFromRow">The optional row to copy from.</param>
    private sealed record InsertRowParameters(int SlideIndex, int ShapeIndex, int? RowIndex, int? CopyFromRow);
}
