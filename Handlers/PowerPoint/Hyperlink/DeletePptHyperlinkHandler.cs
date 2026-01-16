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
        var p = ExtractDeleteParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);
        var shape = PowerPointHelper.GetShape(slide, p.ShapeIndex);
        shape.HyperlinkClick = null;

        if (shape is IAutoShape { TextFrame: not null } autoShape)
            foreach (var paragraph in autoShape.TextFrame.Paragraphs)
            foreach (var portion in paragraph.Portions)
                portion.PortionFormat.HyperlinkClick = null;

        MarkModified(context);

        return Success($"Hyperlink deleted from slide {p.SlideIndex}, shape {p.ShapeIndex}.");
    }

    /// <summary>
    ///     Extracts delete parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted delete parameters.</returns>
    private static DeleteParameters ExtractDeleteParameters(OperationParameters parameters)
    {
        return new DeleteParameters(
            parameters.GetRequired<int>("slideIndex"),
            parameters.GetRequired<int>("shapeIndex")
        );
    }

    /// <summary>
    ///     Record for holding delete hyperlink parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ShapeIndex">The shape index.</param>
    private record DeleteParameters(int SlideIndex, int ShapeIndex);
}
