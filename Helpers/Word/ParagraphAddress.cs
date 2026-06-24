using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Helpers.Word;

/// <summary>
///     A story-relative address of a paragraph. Defaults address the first section's body, so a bare
///     index means "body paragraph N".
/// </summary>
/// <param name="Index">Story-relative paragraph index; -1 = last paragraph in the story.</param>
/// <param name="StoryType">The containing story (see <see cref="StoryTypes" />).</param>
/// <param name="SectionIndex">The section index (used by Body / Header / Footer / TextBox).</param>
/// <param name="HeaderFooterType">For Header / Footer stories: Primary, First, or Even.</param>
/// <param name="ContainerIndex">
///     Instance selector for multi-instance stories (TextBox shape order, Comment id, note order).
/// </param>
/// <param name="Handle">Optional stable handle; when set it takes precedence over the index coordinates.</param>
public sealed record ParagraphAddress(
    int Index,
    string StoryType = StoryTypes.Body,
    int SectionIndex = 0,
    string HeaderFooterType = "Primary",
    int? ContainerIndex = null,
    string? Handle = null)
{
    /// <summary>
    ///     Reads the standard address parameters from an operation-parameter bag, using the supplied
    ///     value for the story-relative paragraph index (whose parameter name varies by operation).
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <param name="index">The already-read story-relative paragraph index.</param>
    /// <returns>The parsed address.</returns>
    public static ParagraphAddress From(OperationParameters parameters, int index)
    {
        return new ParagraphAddress(
            index,
            parameters.GetOptional("storyType", StoryTypes.Body),
            parameters.GetOptional("sectionIndex", 0),
            parameters.GetOptional("headerFooterType", "Primary"),
            parameters.GetOptional<int?>("containerIndex"),
            parameters.GetOptional<string?>("handle"));
    }
}
