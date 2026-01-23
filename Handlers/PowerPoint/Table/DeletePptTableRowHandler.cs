using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Table;

/// <summary>
///     Handler for deleting rows from PowerPoint tables.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class DeletePptTableRowHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "delete_row";

    /// <summary>
    ///     Deletes a row from a table.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: shapeIndex, rowIndex.
    ///     Optional: slideIndex.
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var deleteParams = ExtractDeleteRowParameters(parameters);

        var presentation = context.Document;
        var slide = PptTableHelper.GetSlide(presentation, deleteParams.SlideIndex);
        var table = PptTableHelper.GetTable(slide, deleteParams.ShapeIndex);

        table.Rows.RemoveAt(deleteParams.RowIndex, false);

        MarkModified(context);

        return new SuccessResult { Message = $"Row {deleteParams.RowIndex} deleted." };
    }

    /// <summary>
    ///     Extracts delete row parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted delete row parameters.</returns>
    private static DeleteRowParameters ExtractDeleteRowParameters(OperationParameters parameters)
    {
        return new DeleteRowParameters(
            parameters.GetOptional("slideIndex", 0),
            parameters.GetRequired<int>("shapeIndex"),
            parameters.GetRequired<int>("rowIndex")
        );
    }

    /// <summary>
    ///     Record for holding delete row parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ShapeIndex">The shape index.</param>
    /// <param name="RowIndex">The row index to delete.</param>
    private sealed record DeleteRowParameters(int SlideIndex, int ShapeIndex, int RowIndex);
}
