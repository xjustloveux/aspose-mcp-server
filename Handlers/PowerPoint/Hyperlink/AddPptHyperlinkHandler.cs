using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Hyperlink;

/// <summary>
///     Handler for adding hyperlinks to PowerPoint presentations.
/// </summary>
public class AddPptHyperlinkHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a hyperlink to a shape or specific text portion.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex
    ///     Optional: shapeIndex, text, linkText, url, slideTargetIndex, x, y, width, height
    /// </param>
    /// <returns>Success message with hyperlink details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractHyperlinkParameters(parameters);
        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);

        var autoShape = GetOrCreateAutoShape(slide, p);
        var (hyperlink, linkDescription) = PptHyperlinkHelper.CreateHyperlink(presentation, p.Url, p.SlideTargetIndex);

        var finalDescription = ApplyHyperlink(autoShape, hyperlink, linkDescription, p.Text, p.LinkText);

        MarkModified(context);
        return Success($"Hyperlink added to slide {p.SlideIndex}: {finalDescription}.");
    }

    /// <summary>
    ///     Extracts hyperlink parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted hyperlink parameters.</returns>
    private static HyperlinkParameters ExtractHyperlinkParameters(OperationParameters parameters)
    {
        return new HyperlinkParameters(
            parameters.GetRequired<int>("slideIndex"),
            parameters.GetOptional<int?>("shapeIndex"),
            parameters.GetOptional<string?>("text"),
            parameters.GetOptional<string?>("linkText"),
            parameters.GetOptional<string?>("url"),
            parameters.GetOptional<int?>("slideTargetIndex"),
            parameters.GetOptional<float?>("x") ?? 50f,
            parameters.GetOptional<float?>("y") ?? 50f,
            parameters.GetOptional<float?>("width") ?? 300f,
            parameters.GetOptional<float?>("height") ?? 50f
        );
    }

    /// <summary>
    ///     Gets an existing AutoShape or creates a new one.
    /// </summary>
    /// <param name="slide">The slide.</param>
    /// <param name="p">The hyperlink parameters.</param>
    /// <returns>The AutoShape.</returns>
    private static IAutoShape GetOrCreateAutoShape(ISlide slide, HyperlinkParameters p)
    {
        if (p.ShapeIndex is >= 0 && p.ShapeIndex.Value < slide.Shapes.Count)
        {
            if (slide.Shapes[p.ShapeIndex.Value] is IAutoShape existingAutoShape)
                return existingAutoShape;
            throw new ArgumentException($"Shape at index {p.ShapeIndex.Value} is not an AutoShape");
        }

        return slide.Shapes.AddAutoShape(ShapeType.Rectangle, p.X, p.Y, p.Width, p.Height);
    }

    /// <summary>
    ///     Applies a hyperlink to an AutoShape.
    /// </summary>
    /// <param name="autoShape">The AutoShape to apply the hyperlink to.</param>
    /// <param name="hyperlink">The hyperlink to apply.</param>
    /// <param name="linkDescription">The link description.</param>
    /// <param name="text">The optional text content.</param>
    /// <param name="linkText">The optional link text.</param>
    /// <returns>The final link description.</returns>
    private static string ApplyHyperlink(IAutoShape autoShape, IHyperlink hyperlink, string linkDescription,
        string? text, string? linkText)
    {
        if (!string.IsNullOrEmpty(linkText) && !string.IsNullOrEmpty(text))
            return ApplyPartialTextHyperlink(autoShape, hyperlink, linkDescription, text, linkText);

        autoShape.HyperlinkClick = hyperlink;
        if (!string.IsNullOrEmpty(text) && autoShape.TextFrame != null)
            autoShape.TextFrame.Text = text;

        return linkDescription;
    }

    /// <summary>
    ///     Applies a hyperlink to a specific portion of text.
    /// </summary>
    /// <param name="autoShape">The AutoShape containing the text.</param>
    /// <param name="hyperlink">The hyperlink to apply.</param>
    /// <param name="linkDescription">The link description.</param>
    /// <param name="text">The full text content.</param>
    /// <param name="linkText">The portion of text to apply the hyperlink to.</param>
    /// <returns>The final link description.</returns>
    private static string ApplyPartialTextHyperlink(IAutoShape autoShape, IHyperlink hyperlink,
        string linkDescription, string text, string linkText)
    {
        var linkIndex = text.IndexOf(linkText, StringComparison.Ordinal);
        if (linkIndex < 0)
            throw new ArgumentException($"linkText '{linkText}' not found in text '{text}'");

        autoShape.TextFrame.Paragraphs.Clear();
        var paragraph = new Paragraph();

        if (linkIndex > 0)
            paragraph.Portions.Add(new Portion(text[..linkIndex]));

        var linkPortion = new Portion(linkText) { PortionFormat = { HyperlinkClick = hyperlink } };
        paragraph.Portions.Add(linkPortion);

        var afterIndex = linkIndex + linkText.Length;
        if (afterIndex < text.Length)
            paragraph.Portions.Add(new Portion(text[afterIndex..]));

        autoShape.TextFrame.Paragraphs.Add(paragraph);
        return linkDescription + $" (on text: '{linkText}')";
    }

    /// <summary>
    ///     Record for holding hyperlink parameters extracted from operation parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ShapeIndex">The optional shape index.</param>
    /// <param name="Text">The optional text content.</param>
    /// <param name="LinkText">The optional link text.</param>
    /// <param name="Url">The optional URL.</param>
    /// <param name="SlideTargetIndex">The optional target slide index.</param>
    /// <param name="X">The X coordinate.</param>
    /// <param name="Y">The Y coordinate.</param>
    /// <param name="Width">The width.</param>
    /// <param name="Height">The height.</param>
    private sealed record HyperlinkParameters(
        int SlideIndex,
        int? ShapeIndex,
        string? Text,
        string? LinkText,
        string? Url,
        int? SlideTargetIndex,
        float X,
        float Y,
        float Width,
        float Height);
}
