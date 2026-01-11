using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.DataOperations;

/// <summary>
///     Handler for getting presentation content including text from all shape types.
/// </summary>
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
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var presentation = context.Document;
        List<object> slides = [];

        var slideIndex = 0;
        foreach (var slide in presentation.Slides)
        {
            List<string> textContent = [];
            foreach (var shape in slide.Shapes)
                PowerPointHelper.ExtractTextFromShape(shape, textContent);

            slides.Add(new
            {
                index = slideIndex,
                hidden = slide.Hidden,
                textContent
            });
            slideIndex++;
        }

        var result = new
        {
            totalSlides = presentation.Slides.Count,
            slides
        };

        return JsonResult(result);
    }
}
