using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Word.List;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.List;

/// <summary>
///     Handler for getting list format information from Word documents.
/// </summary>
[ResultType(typeof(GetWordListFormatResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractGetListFormatParameters(parameters);

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

        if (p.ParagraphIndex.HasValue)
        {
            var paraRef = ParagraphResolver.Resolve(doc, ParagraphAddress.From(parameters, p.ParagraphIndex.Value));
            var listInfo = WordListHelper.BuildListFormatSingleResult(paraRef, listItemIndices);

            return listInfo;
        }

        var listParagraphs = paragraphs
            .Where(para => para.ListFormat is { IsListItem: true })
            .ToList();

        if (listParagraphs.Count == 0)
            return new GetWordListFormatResult
            {
                Count = 0,
                ListParagraphs = [],
                Message = "No list paragraphs found"
            };

        List<ListParagraphInfo> listInfos = [];
        var addressingContext = new ParagraphResolver.AddressingContext(doc);
        foreach (var para in listParagraphs)
        {
            var pref = ParagraphResolver.AddressOf(doc, para, addressingContext);
            listInfos.Add(WordListHelper.BuildListParagraphInfo(pref, listItemIndices));
        }

        return new GetWordListFormatResult
        {
            Count = listParagraphs.Count,
            ListParagraphs = listInfos
        };
    }

    private static GetListFormatParameters ExtractGetListFormatParameters(OperationParameters parameters)
    {
        return new GetListFormatParameters(
            parameters.GetOptional<int?>("paragraphIndex"));
    }

    private sealed record GetListFormatParameters(int? ParagraphIndex);
}
