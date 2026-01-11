using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Bookmark;

/// <summary>
///     Handler for adding bookmarks to Word documents.
/// </summary>
public class AddWordBookmarkHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a bookmark to the document at the specified location.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: name
    ///     Optional: text, paragraphIndex
    /// </param>
    /// <returns>Success message with bookmark details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var name = parameters.GetOptional<string?>("name");
        var text = parameters.GetOptional<string?>("text");
        var paragraphIndex = parameters.GetOptional<int?>("paragraphIndex");

        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Bookmark name is required for add operation");

        var doc = context.Document;
        var builder = new DocumentBuilder(doc);

        var existingBookmark = doc.Range.Bookmarks
            .FirstOrDefault(b => b.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        if (existingBookmark != null)
            throw new InvalidOperationException(
                $"Bookmark '{existingBookmark.Name}' already exists (bookmark names are case-insensitive)");

        if (paragraphIndex.HasValue)
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

            if (paragraphs.Count == 0)
            {
                builder.MoveToDocumentEnd();
            }
            else if (paragraphIndex.Value == -1)
            {
                if (paragraphs[0] is WordParagraph firstPara)
                    builder.MoveTo(firstPara);
            }
            else if (paragraphIndex.Value >= 0 && paragraphIndex.Value < paragraphs.Count)
            {
                if (paragraphs[paragraphIndex.Value] is WordParagraph targetPara)
                    builder.MoveTo(targetPara);
                else
                    throw new InvalidOperationException($"Unable to find paragraph at index {paragraphIndex.Value}");
            }
            else
            {
                throw new ArgumentException(
                    $"Paragraph index {paragraphIndex.Value} is out of range (document has {paragraphs.Count} paragraphs)");
            }
        }
        else
        {
            builder.MoveToDocumentEnd();
        }

        builder.StartBookmark(name);
        if (!string.IsNullOrEmpty(text)) builder.Write(text);
        builder.EndBookmark(name);

        MarkModified(context);

        var result = "Bookmark added successfully\n";
        result += $"Bookmark name: {name}\n";
        if (!string.IsNullOrEmpty(text)) result += $"Bookmark text: {text}\n";
        if (paragraphIndex.HasValue)
            result += paragraphIndex.Value == -1
                ? "Insert position: beginning of document"
                : $"Insert position: after paragraph #{paragraphIndex.Value}";
        else
            result += "Insert position: end of document";

        return result;
    }
}
