using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Hyperlink;

/// <summary>
///     Handler for editing hyperlinks in PowerPoint presentations.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractEditParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);
        var shape = PowerPointHelper.GetShape(slide, p.ShapeIndex);

        ApplyHyperlinkChange(shape, p, presentation);

        MarkModified(context);

        return new SuccessResult { Message = $"Hyperlink updated on slide {p.SlideIndex}, shape {p.ShapeIndex}." };
    }

    /// <summary>
    ///     Applies the hyperlink change based on parameters.
    /// </summary>
    /// <param name="shape">The shape to modify.</param>
    /// <param name="p">The edit parameters.</param>
    /// <param name="presentation">The presentation for slide targets.</param>
    /// <exception cref="ArgumentException">Thrown when no valid hyperlink option is provided.</exception>
    private static void ApplyHyperlinkChange(IShape shape, EditParameters p, Presentation presentation)
    {
        if (p.ShouldRemoveHyperlink)
        {
            RemoveHyperlink(shape);
            return;
        }

        if (!string.IsNullOrEmpty(p.Url))
        {
            shape.HyperlinkClick = new Aspose.Slides.Hyperlink(p.Url);
            return;
        }

        if (p.SlideTargetIndex.HasValue)
        {
            SetSlideTargetHyperlink(shape, p.SlideTargetIndex.Value, presentation);
            return;
        }

        throw new ArgumentException("Either url, slideTargetIndex, or removeHyperlink must be provided");
    }

    /// <summary>
    ///     Removes hyperlinks from a shape.
    /// </summary>
    /// <param name="shape">The shape to remove hyperlinks from.</param>
    private static void RemoveHyperlink(IShape shape)
    {
        if (shape is IAutoShape { TextFrame: not null } autoShape)
            foreach (var paragraph in autoShape.TextFrame.Paragraphs)
            foreach (var portion in paragraph.Portions)
                portion.PortionFormat.HyperlinkClick = null;

        shape.HyperlinkClick = null;
    }

    /// <summary>
    ///     Sets a slide target hyperlink on a shape.
    /// </summary>
    /// <param name="shape">The shape to set the hyperlink on.</param>
    /// <param name="slideTargetIndex">The target slide index.</param>
    /// <param name="presentation">The presentation containing the slides.</param>
    /// <exception cref="ArgumentException">Thrown when the slide index is out of range.</exception>
    private static void SetSlideTargetHyperlink(IShape shape, int slideTargetIndex, Presentation presentation)
    {
        if (slideTargetIndex < 0 || slideTargetIndex >= presentation.Slides.Count)
            throw new ArgumentException($"slideTargetIndex must be between 0 and {presentation.Slides.Count - 1}");

        shape.HyperlinkClick = new Aspose.Slides.Hyperlink(presentation.Slides[slideTargetIndex]);
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
    /// <param name="ShouldRemoveHyperlink">Whether to remove the hyperlink.</param>
    private sealed record EditParameters(
        int SlideIndex,
        int ShapeIndex,
        string? Url,
        int? SlideTargetIndex,
        bool ShouldRemoveHyperlink);
}
