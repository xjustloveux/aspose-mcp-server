using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.PowerPoint.DataOperations;

namespace AsposeMcpServer.Handlers.PowerPoint.DataOperations;

/// <summary>
///     Handler for getting presentation content including text from all shape types.
/// </summary>
[ResultType(typeof(GetContentPptResult))]
public class GetContentHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "get_content";

    /// <summary>
    ///     Gets presentation content including text from all slides.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">No parameters required.</param>
    /// <returns>JSON string containing text content extracted from all slides.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var presentation = context.Document;
        List<GetContentSlideInfo> slides = [];

        var slideIndex = 0;
        foreach (var slide in presentation.Slides)
        {
            List<string> textContent = [];
            foreach (var shape in slide.Shapes)
                PowerPointHelper.ExtractTextFromShape(shape, textContent);

            slides.Add(new GetContentSlideInfo
            {
                Index = slideIndex,
                Hidden = slide.Hidden,
                TextContent = textContent
            });
            slideIndex++;
        }

        var result = new GetContentPptResult
        {
            TotalSlides = presentation.Slides.Count,
            Slides = slides
        };

        return result;
    }
}
