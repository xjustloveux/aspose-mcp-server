using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.PowerPoint.Hyperlink;

namespace AsposeMcpServer.Handlers.PowerPoint.Hyperlink;

/// <summary>
///     Handler for getting hyperlink information from PowerPoint presentations.
/// </summary>
[ResultType(typeof(GetHyperlinksPptResult))]
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
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractGetParameters(parameters);

        var presentation = context.Document;

        if (p.SlideIndex.HasValue)
        {
            if (p.SlideIndex.Value < 0 || p.SlideIndex.Value >= presentation.Slides.Count)
                throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
            var slide = presentation.Slides[p.SlideIndex.Value];
            var hyperlinksList = PptHyperlinkHelper.GetHyperlinksFromSlide(presentation, slide);

            var result = new GetHyperlinksPptResult
            {
                SlideIndex = p.SlideIndex.Value,
                Count = hyperlinksList.Count,
                Hyperlinks = hyperlinksList
            };

            return result;
        }
        else
        {
            List<SlideHyperlinksInfo> slidesList = [];
            var totalCount = 0;

            for (var i = 0; i < presentation.Slides.Count; i++)
            {
                var slide = presentation.Slides[i];
                var hyperlinksList = PptHyperlinkHelper.GetHyperlinksFromSlide(presentation, slide);
                totalCount += hyperlinksList.Count;

                slidesList.Add(new SlideHyperlinksInfo
                {
                    SlideIndex = i,
                    Count = hyperlinksList.Count,
                    Hyperlinks = hyperlinksList
                });
            }

            var result = new GetHyperlinksPptResult
            {
                TotalCount = totalCount,
                Slides = slidesList
            };

            return result;
        }
    }

    /// <summary>
    ///     Extracts get parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted get parameters.</returns>
    private static GetParameters ExtractGetParameters(OperationParameters parameters)
    {
        return new GetParameters(parameters.GetOptional<int?>("slideIndex"));
    }

    /// <summary>
    ///     Record for holding get hyperlinks parameters.
    /// </summary>
    /// <param name="SlideIndex">The optional slide index.</param>
    private sealed record GetParameters(int? SlideIndex);
}
