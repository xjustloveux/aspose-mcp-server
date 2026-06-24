using Aspose.Words.Replacing;
using AsposeMcpServer.Helpers.Word;

namespace AsposeMcpServer.Handlers.Word.Text;

/// <summary>
///     Callback to skip field replacement during text replacement operations.
/// </summary>
internal class FieldSkipReplacingCallback : IReplacingCallback
{
    /// <summary>
    ///     Determines whether to replace or skip text replacement based on field context.
    /// </summary>
    /// <param name="args">Replacing arguments containing match information.</param>
    /// <returns>ReplaceAction.Skip if the match begins inside a field, ReplaceAction.Replace otherwise.</returns>
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        return FieldBoundaryHelper.GetEnclosingField(args.MatchNode) != null
            ? ReplaceAction.Skip
            : ReplaceAction.Replace;
    }
}
