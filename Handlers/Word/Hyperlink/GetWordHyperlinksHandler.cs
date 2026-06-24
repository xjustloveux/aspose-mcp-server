using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Word.Hyperlink;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Hyperlink;

/// <summary>
///     Handler for getting hyperlinks from Word documents.
/// </summary>
[ResultType(typeof(GetHyperlinksResult))]
public class GetWordHyperlinksHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all hyperlinks from the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">No parameters required.</param>
    /// <returns>A GetHyperlinksResult containing hyperlink information.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var doc = context.Document;
        var hyperlinkFields = WordHyperlinkHelper.GetAllHyperlinks(doc);

        if (hyperlinkFields.Count == 0)
            return new GetHyperlinksResult
            {
                Count = 0,
                Hyperlinks = [],
                Message = "No hyperlinks found in document"
            };

        List<HyperlinkInfo> hyperlinkList = [];
        var addressingContext = new ParagraphResolver.AddressingContext(doc);
        for (var index = 0; index < hyperlinkFields.Count; index++)
        {
            var hyperlinkField = hyperlinkFields[index];
            var displayText = "";
            var address = "";
            var subAddress = "";
            var tooltip = "";
            ParagraphRef? pref = null;

            try
            {
                displayText = hyperlinkField.Result ?? "";
                address = hyperlinkField.Address ?? "";
                subAddress = hyperlinkField.SubAddress ?? "";
                tooltip = hyperlinkField.ScreenTip ?? "";

                if (hyperlinkField.Start?.ParentNode is WordParagraph para)
                    pref = ParagraphResolver.AddressOf(doc, para, addressingContext);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[WARN] Error reading hyperlink properties: {ex.Message}");
            }

            var addr = pref?.Address;
            var headerFooterType = addr is not null &&
                                   (addr.StoryType == StoryTypes.Header || addr.StoryType == StoryTypes.Footer)
                ? addr.HeaderFooterType
                : null;

            hyperlinkList.Add(new HyperlinkInfo
            {
                Index = index,
                DisplayText = displayText,
                Address = address,
                SubAddress = string.IsNullOrEmpty(subAddress) ? null : subAddress,
                Tooltip = string.IsNullOrEmpty(tooltip) ? null : tooltip,
                ParagraphIndex = addr?.Index,
                StoryType = addr?.StoryType,
                SectionIndex = addr?.SectionIndex,
                HeaderFooterType = headerFooterType,
                ContainerIndex = addr?.ContainerIndex,
                DocumentOrderIndex = pref?.DocumentOrderIndex
            });
        }

        return new GetHyperlinksResult
        {
            Count = hyperlinkFields.Count,
            Hyperlinks = hyperlinkList
        };
    }
}
