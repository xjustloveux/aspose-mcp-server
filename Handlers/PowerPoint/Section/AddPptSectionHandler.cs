using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Section;

/// <summary>
///     Handler for adding sections to PowerPoint presentations.
/// </summary>
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
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var name = parameters.GetRequired<string>("name");
        var slideIndex = parameters.GetRequired<int>("slideIndex");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        presentation.Sections.AddSection(name, slide);

        MarkModified(context);

        return Success($"Section '{name}' added starting at slide {slideIndex}.");
    }
}
