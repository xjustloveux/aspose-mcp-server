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
    /// <exception cref="ArgumentException">Thrown when bookmark name is not provided.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractAddParameters(parameters);

        var doc = context.Document;
        var builder = new DocumentBuilder(doc);

        ValidateBookmarkDoesNotExist(doc, p.Name);
        MoveToInsertPosition(builder, doc, p.ParagraphIndex);

        builder.StartBookmark(p.Name);
        if (!string.IsNullOrEmpty(p.Text)) builder.Write(p.Text);
        builder.EndBookmark(p.Name);

        MarkModified(context);

        return BuildResultMessage(p.Name, p.Text, p.ParagraphIndex);
    }

    /// <summary>
    ///     Validates that a bookmark with the specified name does not already exist.
    /// </summary>
    /// <param name="doc">The Word document to check.</param>
    /// <param name="name">The bookmark name to validate.</param>
    /// <exception cref="InvalidOperationException">Thrown when a bookmark with the same name already exists.</exception>
    private static void ValidateBookmarkDoesNotExist(Document doc, string name)
    {
        var existingBookmark = doc.Range.Bookmarks
            .FirstOrDefault(b => b.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        if (existingBookmark != null)
            throw new InvalidOperationException(
                $"Bookmark '{existingBookmark.Name}' already exists (bookmark names are case-insensitive)");
    }

    /// <summary>
    ///     Moves the document builder to the specified insert position.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="doc">The Word document.</param>
    /// <param name="paragraphIndex">The paragraph index to move to, or null for end of document.</param>
    /// <exception cref="ArgumentException">Thrown when the paragraph index is out of range.</exception>
    private static void MoveToInsertPosition(DocumentBuilder builder, Document doc, int? paragraphIndex)
    {
        if (!paragraphIndex.HasValue)
        {
            builder.MoveToDocumentEnd();
            return;
        }

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        if (paragraphs.Count == 0)
        {
            builder.MoveToDocumentEnd();
            return;
        }

        if (paragraphIndex.Value == -1)
        {
            MoveToParagraph(builder, paragraphs[0], 0);
            return;
        }

        if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
            throw new ArgumentException(
                $"Paragraph index {paragraphIndex.Value} is out of range (document has {paragraphs.Count} paragraphs)");

        MoveToParagraph(builder, paragraphs[paragraphIndex.Value], paragraphIndex.Value);
    }

    /// <summary>
    ///     Moves the document builder to the specified paragraph node.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="node">The paragraph node to move to.</param>
    /// <param name="index">The paragraph index for error reporting.</param>
    /// <exception cref="InvalidOperationException">Thrown when the node is not a valid paragraph.</exception>
    private static void MoveToParagraph(DocumentBuilder builder, Node node, int index)
    {
        if (node is WordParagraph para)
            builder.MoveTo(para);
        else
            throw new InvalidOperationException($"Unable to find paragraph at index {index}");
    }

    /// <summary>
    ///     Builds the result message for a successful bookmark addition.
    /// </summary>
    /// <param name="name">The bookmark name.</param>
    /// <param name="text">The bookmark text content.</param>
    /// <param name="paragraphIndex">The paragraph index where the bookmark was inserted.</param>
    /// <returns>A formatted result message.</returns>
    private static string BuildResultMessage(string name, string? text, int? paragraphIndex)
    {
        var result = "Bookmark added successfully\n";
        result += $"Bookmark name: {name}\n";
        if (!string.IsNullOrEmpty(text)) result += $"Bookmark text: {text}\n";
        result += GetInsertPositionMessage(paragraphIndex);
        return result;
    }

    /// <summary>
    ///     Gets the insert position message based on paragraph index.
    /// </summary>
    /// <param name="paragraphIndex">The paragraph index.</param>
    /// <returns>A message describing the insert position.</returns>
    private static string GetInsertPositionMessage(int? paragraphIndex)
    {
        if (!paragraphIndex.HasValue)
            return "Insert position: end of document";
        return paragraphIndex.Value == -1
            ? "Insert position: beginning of document"
            : $"Insert position: after paragraph #{paragraphIndex.Value}";
    }

    /// <summary>
    ///     Extracts and validates parameters for the add bookmark operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    /// <exception cref="ArgumentException">Thrown when bookmark name is not provided.</exception>
    private static AddParameters ExtractAddParameters(OperationParameters parameters)
    {
        var name = parameters.GetOptional<string?>("name");
        var text = parameters.GetOptional<string?>("text");
        var paragraphIndex = parameters.GetOptional<int?>("paragraphIndex");

        if (string.IsNullOrEmpty(name))
            throw new ArgumentException("Bookmark name is required for add operation");

        return new AddParameters(name, text, paragraphIndex);
    }

    /// <summary>
    ///     Parameters for the add bookmark operation.
    /// </summary>
    /// <param name="Name">The bookmark name.</param>
    /// <param name="Text">The text content for the bookmark.</param>
    /// <param name="ParagraphIndex">The paragraph index where the bookmark should be inserted.</param>
    private record AddParameters(string Name, string? Text, int? ParagraphIndex);
}
