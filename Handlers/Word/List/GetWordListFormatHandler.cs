using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.List;

/// <summary>
///     Handler for getting list format information from Word documents.
/// </summary>
public class GetWordListFormatHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get_format";

    /// <summary>
    ///     Gets list format information for a paragraph or all list paragraphs.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: paragraphIndex (if not provided, returns all list paragraphs)
    /// </param>
    /// <returns>JSON string containing list format information.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var paragraphIndex = parameters.GetOptional<int?>("paragraphIndex");

        var doc = context.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();

        var listItemIndices = new Dictionary<(int listId, int paraIndex), int>();
        var listCounters = new Dictionary<int, int>();
        foreach (var para in paragraphs)
            if (para.ListFormat is { IsListItem: true, List: not null })
            {
                var listId = para.ListFormat.List.ListId;
                listCounters.TryAdd(listId, 0);
                var paraIdx = paragraphs.IndexOf(para);
                listItemIndices[(listId, paraIdx)] = listCounters[listId];
                listCounters[listId]++;
            }

        if (paragraphIndex.HasValue)
        {
            if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
                throw new ArgumentException(
                    $"Paragraph index {paragraphIndex.Value} is out of range (document has {paragraphs.Count} paragraphs)");

            var para = paragraphs[paragraphIndex.Value];
            var listInfo = WordListHelper.BuildListFormatInfo(para, paragraphIndex.Value, listItemIndices);

            return JsonResult(listInfo);
        }

        var listParagraphs = paragraphs
            .Where(p => p.ListFormat is { IsListItem: true })
            .ToList();

        if (listParagraphs.Count == 0)
        {
            var emptyResult = new
            {
                count = 0,
                listParagraphs = Array.Empty<object>(),
                message = "No list paragraphs found"
            };
            return JsonResult(emptyResult);
        }

        List<object> listInfos = [];
        foreach (var para in listParagraphs)
        {
            var paraIndex = paragraphs.IndexOf(para);
            listInfos.Add(WordListHelper.BuildListFormatInfo(para, paraIndex, listItemIndices));
        }

        var result = new
        {
            count = listParagraphs.Count,
            listParagraphs = listInfos
        };

        return JsonResult(result);
    }
}
