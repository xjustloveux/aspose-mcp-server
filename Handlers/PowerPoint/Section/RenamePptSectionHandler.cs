using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Section;

/// <summary>
///     Handler for renaming sections in PowerPoint presentations.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractRenamePptSectionParameters(parameters);
        var presentation = context.Document;
        PowerPointHelper.ValidateCollectionIndex(p.SectionIndex, presentation.Sections.Count, "section");

        presentation.Sections[p.SectionIndex].Name = p.NewName;

        MarkModified(context);

        return new SuccessResult { Message = $"Section {p.SectionIndex} renamed to '{p.NewName}'." };
    }

    private static RenamePptSectionParameters ExtractRenamePptSectionParameters(OperationParameters parameters)
    {
        return new RenamePptSectionParameters(
            parameters.GetRequired<int>("sectionIndex"),
            parameters.GetRequired<string>("newName"));
    }

    private sealed record RenamePptSectionParameters(int SectionIndex, string NewName);
}
