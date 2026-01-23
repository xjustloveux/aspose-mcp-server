using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Bookmark;

/// <summary>
///     Handler for getting bookmark location information in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class GotoWordBookmarkHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "goto";

    /// <summary>
    ///     Gets location information for a specific bookmark.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: name
    /// </param>
    /// <returns>A message containing the bookmark's location information.</returns>
    /// <exception cref="ArgumentException">Thrown when bookmark name is not provided or bookmark is not found.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractGotoParameters(parameters);

        var doc = context.Document;

        var bookmark = doc.Range.Bookmarks[p.Name];
        if (bookmark == null)
            throw new ArgumentException(
                $"Bookmark '{p.Name}' not found. Use get operation to view available bookmarks");

        var bookmarkText = bookmark.Text;
        var bookmarkRange = bookmark.BookmarkStart?.ParentNode as WordParagraph;

        var paragraphIndex = -1;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        for (var i = 0; i < paragraphs.Count; i++)
            if (paragraphs[i] == bookmarkRange)
            {
                paragraphIndex = i;
                break;
            }

        var message = "Bookmark location information\n";
        message += $"Bookmark name: {p.Name}\n";
        message += $"Bookmark text: {bookmarkText}\n";
        if (paragraphIndex >= 0) message += $"Paragraph index: {paragraphIndex}\n";
        message += $"Bookmark range length: {bookmarkText.Length} characters";

        if (bookmarkRange != null)
        {
            var paraText = bookmarkRange.GetText().Trim();
            if (paraText.Length > 100) paraText = paraText[..100] + "...";
            message += $"\nParagraph content: {paraText}";
        }

        return new SuccessResult { Message = message };
    }

    /// <summary>
    ///     Extracts and validates parameters for the goto bookmark operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    /// <exception cref="ArgumentException">Thrown when bookmark name is not provided.</exception>
    private static GotoParameters ExtractGotoParameters(OperationParameters parameters)
    {
        var name = parameters.GetOptional<string?>("name");

        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Bookmark name is required for goto operation");

        return new GotoParameters(name);
    }

    /// <summary>
    ///     Parameters for the goto bookmark operation.
    /// </summary>
    /// <param name="Name">The bookmark name to navigate to.</param>
    private sealed record GotoParameters(string Name);
}
