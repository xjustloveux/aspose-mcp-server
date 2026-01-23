using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Section;

/// <summary>
///     Handler for deleting sections from PowerPoint presentations.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractDeletePptSectionParameters(parameters);
        var presentation = context.Document;
        PowerPointHelper.ValidateCollectionIndex(p.SectionIndex, presentation.Sections.Count, "section");

        var section = presentation.Sections[p.SectionIndex];
        if (p.KeepSlides)
            presentation.Sections.RemoveSection(section);
        else
            presentation.Sections.RemoveSectionWithSlides(section);

        MarkModified(context);

        return new SuccessResult { Message = $"Section {p.SectionIndex} removed (keep slides: {p.KeepSlides})." };
    }

    private static DeletePptSectionParameters ExtractDeletePptSectionParameters(OperationParameters parameters)
    {
        return new DeletePptSectionParameters(
            parameters.GetRequired<int>("sectionIndex"),
            parameters.GetOptional("keepSlides", true));
    }

    private sealed record DeletePptSectionParameters(int SectionIndex, bool KeepSlides);
}
