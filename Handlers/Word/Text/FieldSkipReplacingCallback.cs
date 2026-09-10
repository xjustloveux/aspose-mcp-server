using Aspose.Words;
using Aspose.Words.Replacing;
using AsposeMcpServer.Helpers.Word;

namespace AsposeMcpServer.Handlers.Word.Text;

/// <summary>
///     Callback to skip field replacement during text replacement operations.
/// </summary>
internal class FieldSkipReplacingCallback : IReplacingCallback
{
    /// <summary>The document the current index describes.</summary>
    private Document? _document;

    /// <summary>
    ///     Where the document's fields are. Kept between matches, because replacing text over a
    ///     large document calls this once per match and numbering the document each time would be
    ///     quadratic — and rebuilt whenever a match arrives that the index does not know, since a
    ///     replacement splits runs and the new ones were never numbered.
    /// </summary>
    private FieldBoundaryHelper.FieldExtents? _extents;

    /// <summary>
    ///     Determines whether to replace or skip text replacement based on field context.
    /// </summary>
    /// <param name="args">Replacing arguments containing match information.</param>
    /// <returns>ReplaceAction.Skip if the match begins inside a field, ReplaceAction.Replace otherwise.</returns>
    public ReplaceAction Replacing(ReplacingArgs args)
    {
        if (args.MatchNode.Document is not Document document) return ReplaceAction.Replace;

        if (_extents == null || !ReferenceEquals(_document, document) || !_extents.Knows(args.MatchNode))
        {
            _document = document;
            _extents = FieldBoundaryHelper.FieldExtents.Of(document);
        }

        return _extents.EnclosingField(args.MatchNode) != null
            ? ReplaceAction.Skip
            : ReplaceAction.Replace;
    }
}
