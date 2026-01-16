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
        var p = ExtractRestartNumberingParameters(parameters);

        var doc = context.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();

        if (p.ParagraphIndex < 0 || p.ParagraphIndex >= paragraphs.Count)
            throw new ArgumentException(
                $"Paragraph index {p.ParagraphIndex} is out of range (document has {paragraphs.Count} paragraphs)");

        var para = paragraphs[p.ParagraphIndex];

        if (!para.ListFormat.IsListItem)
            throw new ArgumentException(
                $"Paragraph at index {p.ParagraphIndex} is not a list item. Use get_format operation to find list item paragraphs.");

        var originalList = para.ListFormat.List;
        if (originalList == null)
            throw new InvalidOperationException("Unable to access list for this paragraph");

        var newList = doc.Lists.AddCopy(originalList);
        var level = para.ListFormat.ListLevelNumber;

        newList.ListLevels[level].StartAt = p.StartAt;

        var applyCount = 0;
        for (var i = p.ParagraphIndex; i < paragraphs.Count; i++)
        {
            var currentPara = paragraphs[i];
            if (currentPara.ListFormat.IsListItem && currentPara.ListFormat.List?.ListId == originalList.ListId)
            {
                currentPara.ListFormat.List = newList;
                applyCount++;
            }
            else if (i > p.ParagraphIndex && !currentPara.ListFormat.IsListItem)
            {
                break;
            }
        }

        MarkModified(context);

        var result = "List numbering restarted successfully\n";
        result += $"Paragraph index: {p.ParagraphIndex}\n";
        result += $"Start at: {p.StartAt}\n";
        result += $"Paragraphs affected: {applyCount}\n";
        result += $"New list ID: {newList.ListId}";

        return Success(result);
    }

    private static RestartNumberingParameters ExtractRestartNumberingParameters(OperationParameters parameters)
    {
        return new RestartNumberingParameters(
            parameters.GetRequired<int>("paragraphIndex"),
            parameters.GetOptional("startAt", 1));
    }

    private record RestartNumberingParameters(
        int ParagraphIndex,
        int StartAt);
}
