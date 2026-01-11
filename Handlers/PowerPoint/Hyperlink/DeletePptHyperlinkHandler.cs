using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Hyperlink;

/// <summary>
///     Handler for deleting hyperlinks from PowerPoint presentations.
/// </summary>
public class DeletePptHyperlinkHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a hyperlink from a shape.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex, shapeIndex
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetRequired<int>("slideIndex");
        var shapeIndex = parameters.GetRequired<int>("shapeIndex");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex);
        shape.HyperlinkClick = null;

        if (shape is IAutoShape { TextFrame: not null } autoShape)
            foreach (var paragraph in autoShape.TextFrame.Paragraphs)
            foreach (var portion in paragraph.Portions)
                portion.PortionFormat.HyperlinkClick = null;

        MarkModified(context);

        return Success($"Hyperlink deleted from slide {slideIndex}, shape {shapeIndex}.");
    }
}
