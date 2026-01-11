using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Section;

/// <summary>
///     Handler for renaming sections in PowerPoint presentations.
/// </summary>
public class RenamePptSectionHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "rename";

    /// <summary>
    ///     Renames a section in the presentation.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: sectionIndex, newName
    /// </param>
    /// <returns>Success message with rename details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var sectionIndex = parameters.GetRequired<int>("sectionIndex");
        var newName = parameters.GetRequired<string>("newName");

        var presentation = context.Document;
        PowerPointHelper.ValidateCollectionIndex(sectionIndex, presentation.Sections.Count, "section");

        presentation.Sections[sectionIndex].Name = newName;

        MarkModified(context);

        return Success($"Section {sectionIndex} renamed to '{newName}'.");
    }
}
