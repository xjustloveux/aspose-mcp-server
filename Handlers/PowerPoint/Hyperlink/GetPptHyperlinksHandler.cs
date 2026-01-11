using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Hyperlink;

/// <summary>
///     Handler for getting hyperlink information from PowerPoint presentations.
/// </summary>
public class GetPptHyperlinksHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets hyperlink information for a slide or all slides.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Optional: slideIndex
    /// </param>
    /// <returns>JSON string containing the hyperlink information.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetOptional<int?>("slideIndex");

        var presentation = context.Document;

        if (slideIndex.HasValue)
        {
            if (slideIndex.Value < 0 || slideIndex.Value >= presentation.Slides.Count)
                throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
            var slide = presentation.Slides[slideIndex.Value];
            var hyperlinksList = PptHyperlinkHelper.GetHyperlinksFromSlide(presentation, slide);

            var result = new
            {
                slideIndex = slideIndex.Value,
                count = hyperlinksList.Count,
                hyperlinks = hyperlinksList
            };

            return JsonResult(result);
        }
        else
        {
            List<object> slidesList = [];
            var totalCount = 0;

            for (var i = 0; i < presentation.Slides.Count; i++)
            {
                var slide = presentation.Slides[i];
                var hyperlinksList = PptHyperlinkHelper.GetHyperlinksFromSlide(presentation, slide);
                totalCount += hyperlinksList.Count;

                slidesList.Add(new
                {
                    slideIndex = i,
                    count = hyperlinksList.Count,
                    hyperlinks = hyperlinksList
                });
            }

            var result = new
            {
                totalCount,
                slides = slidesList
            };

            return JsonResult(result);
        }
    }
}
