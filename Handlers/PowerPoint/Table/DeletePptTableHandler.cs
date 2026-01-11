using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Table;

/// <summary>
///     Handler for deleting tables from PowerPoint presentations.
/// </summary>
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
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetOptional("slideIndex", 0);
        var shapeIndex = parameters.GetRequired<int>("shapeIndex");

        var presentation = context.Document;
        var slide = PptTableHelper.GetSlide(presentation, slideIndex);

        _ = PptTableHelper.GetTable(slide, shapeIndex);

        slide.Shapes.RemoveAt(shapeIndex);

        MarkModified(context);

        return Success($"Table deleted from slide {slideIndex}.");
    }
}
