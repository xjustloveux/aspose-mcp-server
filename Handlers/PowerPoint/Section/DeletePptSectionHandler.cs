using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Section;

/// <summary>
///     Handler for deleting sections from PowerPoint presentations.
/// </summary>
public class DeletePptSectionHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a section from the presentation.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: sectionIndex
    ///     Optional: keepSlides (default: true)
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var sectionIndex = parameters.GetRequired<int>("sectionIndex");
        var keepSlides = parameters.GetOptional("keepSlides", true);

        var presentation = context.Document;
        PowerPointHelper.ValidateCollectionIndex(sectionIndex, presentation.Sections.Count, "section");

        var section = presentation.Sections[sectionIndex];
        if (keepSlides)
            presentation.Sections.RemoveSection(section);
        else
            presentation.Sections.RemoveSectionWithSlides(section);

        MarkModified(context);

        return Success($"Section {sectionIndex} removed (keep slides: {keepSlides}).");
    }
}
