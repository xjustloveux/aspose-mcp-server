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
        var p = ExtractEditParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);
        var shape = PowerPointHelper.GetShape(slide, p.ShapeIndex);

        if (p.RemoveHyperlink)
        {
            if (shape is IAutoShape { TextFrame: not null } autoShape)
                foreach (var paragraph in autoShape.TextFrame.Paragraphs)
                foreach (var portion in paragraph.Portions)
                    portion.PortionFormat.HyperlinkClick = null;

            shape.HyperlinkClick = null;
        }
        else if (!string.IsNullOrEmpty(p.Url))
        {
            shape.HyperlinkClick = new Aspose.Slides.Hyperlink(p.Url);
        }
        else if (p.SlideTargetIndex.HasValue)
        {
            if (p.SlideTargetIndex.Value < 0 || p.SlideTargetIndex.Value >= presentation.Slides.Count)
                throw new ArgumentException(
                    $"slideTargetIndex must be between 0 and {presentation.Slides.Count - 1}");
            shape.HyperlinkClick = new Aspose.Slides.Hyperlink(presentation.Slides[p.SlideTargetIndex.Value]);
        }
        else
        {
            throw new ArgumentException("Either url, slideTargetIndex, or removeHyperlink must be provided");
        }

        MarkModified(context);

        return Success($"Hyperlink updated on slide {p.SlideIndex}, shape {p.ShapeIndex}.");
    }

    /// <summary>
    ///     Extracts edit parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted edit parameters.</returns>
    private static EditParameters ExtractEditParameters(OperationParameters parameters)
    {
        return new EditParameters(
            parameters.GetRequired<int>("slideIndex"),
            parameters.GetRequired<int>("shapeIndex"),
            parameters.GetOptional<string?>("url"),
            parameters.GetOptional<int?>("slideTargetIndex"),
            parameters.GetOptional<bool?>("removeHyperlink") ?? false
        );
    }

    /// <summary>
    ///     Record for holding edit hyperlink parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ShapeIndex">The shape index.</param>
    /// <param name="Url">The optional URL.</param>
    /// <param name="SlideTargetIndex">The optional target slide index.</param>
    /// <param name="RemoveHyperlink">Whether to remove the hyperlink.</param>
    private record EditParameters(
        int SlideIndex,
        int ShapeIndex,
        string? Url,
        int? SlideTargetIndex,
        bool RemoveHyperlink);
}
