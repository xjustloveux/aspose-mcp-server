using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Table;

/// <summary>
///     Handler for deleting tables from PowerPoint presentations.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class DeletePptTableHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a table from a slide.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: shapeIndex.
    ///     Optional: slideIndex (default: 0).
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var deleteParams = ExtractDeleteParameters(parameters);

        var presentation = context.Document;
        var slide = PptTableHelper.GetSlide(presentation, deleteParams.SlideIndex);

        _ = PptTableHelper.GetTable(slide, deleteParams.ShapeIndex);

        slide.Shapes.RemoveAt(deleteParams.ShapeIndex);

        MarkModified(context);

        return new SuccessResult { Message = $"Table deleted from slide {deleteParams.SlideIndex}." };
    }

    /// <summary>
    ///     Extracts delete parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted delete parameters.</returns>
    private static DeleteParameters ExtractDeleteParameters(OperationParameters parameters)
    {
        return new DeleteParameters(
            parameters.GetOptional("slideIndex", 0),
            parameters.GetRequired<int>("shapeIndex")
        );
    }

    /// <summary>
    ///     Record for holding delete table parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ShapeIndex">The shape index.</param>
    private sealed record DeleteParameters(int SlideIndex, int ShapeIndex);
}
