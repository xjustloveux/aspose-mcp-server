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
        var slideIndex = parameters.GetRequired<int>("slideIndex");
        var shapeIndex = parameters.GetOptional<int?>("shapeIndex");
        var text = parameters.GetOptional<string?>("text");
        var linkText = parameters.GetOptional<string?>("linkText");
        var url = parameters.GetOptional<string?>("url");
        var slideTargetIndex = parameters.GetOptional<int?>("slideTargetIndex");
        var x = parameters.GetOptional<float?>("x") ?? 50f;
        var y = parameters.GetOptional<float?>("y") ?? 50f;
        var width = parameters.GetOptional<float?>("width") ?? 300f;
        var height = parameters.GetOptional<float?>("height") ?? 50f;

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        IAutoShape autoShape;

        if (shapeIndex is >= 0 && shapeIndex.Value < slide.Shapes.Count)
        {
            if (slide.Shapes[shapeIndex.Value] is IAutoShape existingAutoShape)
                autoShape = existingAutoShape;
            else
                throw new ArgumentException($"Shape at index {shapeIndex.Value} is not an AutoShape");
        }
        else
        {
            autoShape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, x, y, width, height);
        }

        var (hyperlink, linkDescription) = PptHyperlinkHelper.CreateHyperlink(presentation, url, slideTargetIndex);

        if (!string.IsNullOrEmpty(linkText) && !string.IsNullOrEmpty(text))
        {
            var linkIndex = text.IndexOf(linkText, StringComparison.Ordinal);
            if (linkIndex < 0)
                throw new ArgumentException($"linkText '{linkText}' not found in text '{text}'");

            autoShape.TextFrame.Paragraphs.Clear();
            var paragraph = new Paragraph();

            if (linkIndex > 0)
            {
                var beforePortion = new Portion(text[..linkIndex]);
                paragraph.Portions.Add(beforePortion);
            }

            var linkPortion = new Portion(linkText)
            {
                PortionFormat = { HyperlinkClick = hyperlink }
            };
            paragraph.Portions.Add(linkPortion);

            var afterIndex = linkIndex + linkText.Length;
            if (afterIndex < text.Length)
            {
                var afterPortion = new Portion(text[afterIndex..]);
                paragraph.Portions.Add(afterPortion);
            }

            autoShape.TextFrame.Paragraphs.Add(paragraph);
            linkDescription += $" (on text: '{linkText}')";
        }
        else
        {
            autoShape.HyperlinkClick = hyperlink;

            if (!string.IsNullOrEmpty(text) && autoShape.TextFrame != null)
                autoShape.TextFrame.Text = text;
        }

        MarkModified(context);

        return Success($"Hyperlink added to slide {slideIndex}: {linkDescription}.");
    }
}
