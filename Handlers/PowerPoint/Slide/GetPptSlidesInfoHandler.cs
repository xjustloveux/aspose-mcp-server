using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.PowerPoint.Slide;

namespace AsposeMcpServer.Handlers.PowerPoint.Slide;

/// <summary>
///     Handler for getting slides information from PowerPoint presentations.
/// </summary>
[ResultType(typeof(GetSlidesInfoResult))]
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
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        _ = parameters;

        var presentation = context.Document;

        List<GetSlideInfoItem> slidesList = [];
        for (var i = 0; i < presentation.Slides.Count; i++)
        {
            var slide = presentation.Slides[i];
            var title = slide.Shapes.FirstOrDefault(s =>
                s.Placeholder?.Type == PlaceholderType.Title) as IAutoShape;
            var titleText = title?.TextFrame?.Text ?? "(no title)";
            var notes = slide.NotesSlideManager.NotesSlide?.NotesTextFrame?.Text;

            slidesList.Add(new GetSlideInfoItem
            {
                Index = i,
                Title = titleText,
                LayoutType = slide.LayoutSlide.LayoutType.ToString(),
                LayoutName = slide.LayoutSlide.Name,
                ShapesCount = slide.Shapes.Count,
                HasSpeakerNotes = !string.IsNullOrWhiteSpace(notes),
                Hidden = slide.Hidden
            });
        }

        var layoutsList = presentation.LayoutSlides
            .Select((ls, idx) => new GetSlideLayoutInfo
                { Index = idx, Name = ls.Name, Type = ls.LayoutType.ToString() })
            .ToList();

        var result = new GetSlidesInfoResult
        {
            Count = presentation.Slides.Count,
            Slides = slidesList,
            AvailableLayouts = layoutsList
        };

        return result;
    }
}
