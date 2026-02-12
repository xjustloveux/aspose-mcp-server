using Aspose.Words;
using Aspose.Words.Replacing;

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
    /// <returns>ReplaceAction.Skip if inside a field, ReplaceAction.Replace otherwise.</returns>
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        if (args.MatchNode.GetAncestor(NodeType.FieldStart) != null ||
            args.MatchNode.GetAncestor(NodeType.FieldSeparator) != null ||
            args.MatchNode.GetAncestor(NodeType.FieldEnd) != null)
            return ReplaceAction.Skip;
        return ReplaceAction.Replace;
    }
}
