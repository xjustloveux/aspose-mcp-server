using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Section;

/// <summary>
///     Handler for getting sections from PowerPoint presentations.
/// </summary>
public class GetPptSectionsHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all sections from the presentation.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">No parameters required.</param>
    /// <returns>JSON string containing section information.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var presentation = context.Document;

        if (presentation.Sections.Count == 0)
            return JsonResult(new
            {
                count = 0,
                sections = Array.Empty<object>(),
                message = "No sections found"
            });

        List<object> sectionsList = [];
        for (var i = 0; i < presentation.Sections.Count; i++)
        {
            var sec = presentation.Sections[i];
            var startSlideIndex = sec.StartedFromSlide != null
                ? presentation.Slides.IndexOf(sec.StartedFromSlide)
                : -1;
            sectionsList.Add(new
            {
                index = i,
                name = sec.Name,
                startSlideIndex,
                slideCount = sec.GetSlidesListOfSection().Count
            });
        }

        return JsonResult(new
        {
            count = presentation.Sections.Count,
            sections = sectionsList
        });
    }
}
