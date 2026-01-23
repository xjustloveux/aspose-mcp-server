using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.PowerPoint.Section;

namespace AsposeMcpServer.Handlers.PowerPoint.Section;

/// <summary>
///     Handler for getting sections from PowerPoint presentations.
/// </summary>
[ResultType(typeof(GetSectionsResult))]
public class GetPptSectionsHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all sections from the presentation.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">No parameters required.</param>
    /// <returns>GetSectionsResult containing section information.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        _ = parameters;

        var presentation = context.Document;

        if (presentation.Sections.Count == 0)
            return new GetSectionsResult
            {
                Count = 0,
                Sections = [],
                Message = "No sections found"
            };

        List<SectionInfo> sectionsList = [];
        for (var i = 0; i < presentation.Sections.Count; i++)
        {
            var sec = presentation.Sections[i];
            var startSlideIndex = sec.StartedFromSlide != null
                ? presentation.Slides.IndexOf(sec.StartedFromSlide)
                : -1;
            sectionsList.Add(new SectionInfo
            {
                Index = i,
                Name = sec.Name,
                StartSlideIndex = startSlideIndex,
                SlideCount = sec.GetSlidesListOfSection().Count
            });
        }

        return new GetSectionsResult
        {
            Count = presentation.Sections.Count,
            Sections = sectionsList
        };
    }
}
