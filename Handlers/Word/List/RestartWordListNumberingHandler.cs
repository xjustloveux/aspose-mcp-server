using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.List;

/// <summary>
///     Handler for restarting list numbering in Word documents.
/// </summary>
public class RestartWordListNumberingHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "restart_numbering";

    /// <summary>
    ///     Restarts list numbering from a specified value at the given paragraph.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: paragraphIndex.
    ///     Optional: startAt (default: 1)
    /// </param>
    /// <returns>Success message with restart details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var paragraphIndex = parameters.GetRequired<int>("paragraphIndex");
        var startAt = parameters.GetOptional("startAt", 1);

        var doc = context.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();

        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
            throw new ArgumentException(
                $"Paragraph index {paragraphIndex} is out of range (document has {paragraphs.Count} paragraphs)");

        var para = paragraphs[paragraphIndex];

        if (!para.ListFormat.IsListItem)
            throw new ArgumentException(
                $"Paragraph at index {paragraphIndex} is not a list item. Use get_format operation to find list item paragraphs.");

        var originalList = para.ListFormat.List;
        if (originalList == null)
            throw new InvalidOperationException("Unable to access list for this paragraph");

        var newList = doc.Lists.AddCopy(originalList);
        var level = para.ListFormat.ListLevelNumber;

        newList.ListLevels[level].StartAt = startAt;

        var applyCount = 0;
        for (var i = paragraphIndex; i < paragraphs.Count; i++)
        {
            var p = paragraphs[i];
            if (p.ListFormat.IsListItem && p.ListFormat.List?.ListId == originalList.ListId)
            {
                p.ListFormat.List = newList;
                applyCount++;
            }
            else if (i > paragraphIndex && !p.ListFormat.IsListItem)
            {
                break;
            }
        }

        MarkModified(context);

        var result = "List numbering restarted successfully\n";
        result += $"Paragraph index: {paragraphIndex}\n";
        result += $"Start at: {startAt}\n";
        result += $"Paragraphs affected: {applyCount}\n";
        result += $"New list ID: {newList.ListId}";

        return Success(result);
    }
}
