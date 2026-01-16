using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Slide;

/// <summary>
///     Handler for getting slides information from PowerPoint presentations.
/// </summary>
public class GetPptSlidesInfoHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "get_info";

    /// <summary>
    ///     Gets information about all slides in the presentation.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">No required parameters.</param>
    /// <returns>JSON string containing slides information including count, slide details, and available layouts.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        _ = ExtractGetPptSlidesInfoParameters(parameters);

        var presentation = context.Document;

        List<object> slidesList = [];
        for (var i = 0; i < presentation.Slides.Count; i++)
        {
            var slide = presentation.Slides[i];
            var title = slide.Shapes.FirstOrDefault(s =>
                s.Placeholder?.Type == PlaceholderType.Title) as IAutoShape;
            var titleText = title?.TextFrame?.Text ?? "(no title)";
            var notes = slide.NotesSlideManager.NotesSlide?.NotesTextFrame?.Text;

            slidesList.Add(new
            {
                index = i,
                title = titleText,
                layoutType = slide.LayoutSlide.LayoutType.ToString(),
                layoutName = slide.LayoutSlide.Name,
                shapesCount = slide.Shapes.Count,
                hasSpeakerNotes = !string.IsNullOrWhiteSpace(notes),
                hidden = slide.Hidden
            });
        }

        var layoutsList = presentation.LayoutSlides
            .Select((ls, idx) => new { index = idx, name = ls.Name, type = ls.LayoutType.ToString() })
            .ToList();

        var result = new
        {
            count = presentation.Slides.Count,
            slides = slidesList,
            availableLayouts = layoutsList
        };

        return JsonResult(result);
    }

    /// <summary>
    ///     Extracts parameters for get slides info operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static GetPptSlidesInfoParameters ExtractGetPptSlidesInfoParameters(OperationParameters parameters)
    {
        _ = parameters;
        return new GetPptSlidesInfoParameters();
    }

    /// <summary>
    ///     Parameters for get slides info operation.
    /// </summary>
    private record GetPptSlidesInfoParameters;
}
