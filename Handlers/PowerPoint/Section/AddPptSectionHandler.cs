using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Section;

/// <summary>
///     Handler for adding sections to PowerPoint presentations.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddPptSectionHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a section to the presentation.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: name, slideIndex
    /// </param>
    /// <returns>Success message with section creation details.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractAddPptSectionParameters(parameters);
        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);
        presentation.Sections.AddSection(p.Name, slide);

        MarkModified(context);

        return new SuccessResult { Message = $"Section '{p.Name}' added starting at slide {p.SlideIndex}." };
    }

    private static AddPptSectionParameters ExtractAddPptSectionParameters(OperationParameters parameters)
    {
        return new AddPptSectionParameters(
            parameters.GetRequired<string>("name"),
            parameters.GetRequired<int>("slideIndex"));
    }

    private sealed record AddPptSectionParameters(string Name, int SlideIndex);
}
