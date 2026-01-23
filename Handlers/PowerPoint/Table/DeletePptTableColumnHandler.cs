using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Table;

/// <summary>
///     Handler for deleting columns from PowerPoint tables.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class DeletePptTableColumnHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "delete_column";

    /// <summary>
    ///     Deletes a column from a table.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: shapeIndex, columnIndex.
    ///     Optional: slideIndex.
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var deleteParams = ExtractDeleteColumnParameters(parameters);

        var presentation = context.Document;
        var slide = PptTableHelper.GetSlide(presentation, deleteParams.SlideIndex);
        var table = PptTableHelper.GetTable(slide, deleteParams.ShapeIndex);

        table.Columns.RemoveAt(deleteParams.ColumnIndex, false);

        MarkModified(context);

        return new SuccessResult { Message = $"Column {deleteParams.ColumnIndex} deleted." };
    }

    /// <summary>
    ///     Extracts delete column parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted delete column parameters.</returns>
    private static DeleteColumnParameters ExtractDeleteColumnParameters(OperationParameters parameters)
    {
        return new DeleteColumnParameters(
            parameters.GetOptional("slideIndex", 0),
            parameters.GetRequired<int>("shapeIndex"),
            parameters.GetRequired<int>("columnIndex")
        );
    }

    /// <summary>
    ///     Record for holding delete column parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ShapeIndex">The shape index.</param>
    /// <param name="ColumnIndex">The column index to delete.</param>
    private sealed record DeleteColumnParameters(int SlideIndex, int ShapeIndex, int ColumnIndex);
}
