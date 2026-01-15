using System.Text.Json;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Hyperlink;

/// <summary>
///     Handler for getting hyperlinks from Word documents.
/// </summary>
public class GetWordHyperlinksHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all hyperlinks from the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">No parameters required.</param>
    /// <returns>A JSON string containing hyperlink information.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var doc = context.Document;
        var hyperlinkFields = WordHyperlinkHelper.GetAllHyperlinks(doc);

        if (hyperlinkFields.Count == 0)
            return JsonSerializer.Serialize(new
                { count = 0, hyperlinks = Array.Empty<object>(), message = "No hyperlinks found in document" });

        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();

        List<object> hyperlinkList = [];
        for (var index = 0; index < hyperlinkFields.Count; index++)
        {
            var hyperlinkField = hyperlinkFields[index];
            var displayText = "";
            var address = "";
            var subAddress = "";
            var tooltip = "";
            int? paragraphIndexValue = null;

            try
            {
                displayText = hyperlinkField.Result ?? "";
                address = hyperlinkField.Address ?? "";
                subAddress = hyperlinkField.SubAddress ?? "";
                tooltip = hyperlinkField.ScreenTip ?? "";

                var fieldStart = hyperlinkField.Start;
                if (fieldStart?.ParentNode is WordParagraph para)
                {
                    paragraphIndexValue = paragraphs.IndexOf(para);
                    if (paragraphIndexValue == -1) paragraphIndexValue = null;
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[WARN] Error reading hyperlink properties: {ex.Message}");
            }

            hyperlinkList.Add(new
            {
                index,
                displayText,
                address,
                subAddress = string.IsNullOrEmpty(subAddress) ? null : subAddress,
                tooltip = string.IsNullOrEmpty(tooltip) ? null : tooltip,
                paragraphIndex = paragraphIndexValue
            });
        }

        var result = new
        {
            count = hyperlinkFields.Count,
            hyperlinks = hyperlinkList
        };

        return JsonSerializer.Serialize(result, JsonDefaults.Indented);
    }
}
