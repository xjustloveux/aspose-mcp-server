using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Hyperlink;

/// <summary>
///     Handler for editing hyperlinks in PowerPoint presentations.
/// </summary>
public class EditPptHyperlinkHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits an existing hyperlink.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex, shapeIndex
    ///     Optional: url, slideTargetIndex, removeHyperlink
    /// </param>
    /// <returns>Success message with update details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetRequired<int>("slideIndex");
        var shapeIndex = parameters.GetRequired<int>("shapeIndex");
        var url = parameters.GetOptional<string?>("url");
        var slideTargetIndex = parameters.GetOptional<int?>("slideTargetIndex");
        var removeHyperlink = parameters.GetOptional<bool?>("removeHyperlink") ?? false;

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex);

        if (removeHyperlink)
        {
            if (shape is IAutoShape { TextFrame: not null } autoShape)
                foreach (var paragraph in autoShape.TextFrame.Paragraphs)
                foreach (var portion in paragraph.Portions)
                    portion.PortionFormat.HyperlinkClick = null;

            shape.HyperlinkClick = null;
        }
        else if (!string.IsNullOrEmpty(url))
        {
            shape.HyperlinkClick = new Aspose.Slides.Hyperlink(url);
        }
        else if (slideTargetIndex.HasValue)
        {
            if (slideTargetIndex.Value < 0 || slideTargetIndex.Value >= presentation.Slides.Count)
                throw new ArgumentException(
                    $"slideTargetIndex must be between 0 and {presentation.Slides.Count - 1}");
            shape.HyperlinkClick = new Aspose.Slides.Hyperlink(presentation.Slides[slideTargetIndex.Value]);
        }
        else
        {
            throw new ArgumentException("Either url, slideTargetIndex, or removeHyperlink must be provided");
        }

        MarkModified(context);

        return Success($"Hyperlink updated on slide {slideIndex}, shape {shapeIndex}.");
    }
}
