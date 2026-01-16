using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Hyperlink;

/// <summary>
///     Handler for adding hyperlinks to Word documents.
/// </summary>
public class AddWordHyperlinkHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a hyperlink to the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: text, either url or subAddress
    ///     Optional: paragraphIndex, tooltip
    /// </param>
    /// <returns>Success message with hyperlink details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var text = parameters.GetOptional<string?>("text");
        var url = parameters.GetOptional<string?>("url");
        var subAddress = parameters.GetOptional<string?>("subAddress");
        var paragraphIndex = parameters.GetOptional<int?>("paragraphIndex");
        var tooltip = parameters.GetOptional<string?>("tooltip");

        if (string.IsNullOrEmpty(url) && string.IsNullOrEmpty(subAddress))
            throw new ArgumentException("Either 'url' or 'subAddress' must be provided for add operation");

        ValidateUrlIfPresent(url);

        var doc = context.Document;
        var builder = new DocumentBuilder(doc);

        PositionBuilder(doc, builder, paragraphIndex);
        InsertHyperlink(builder, text, url, subAddress);
        ApplyHyperlinkSettings(doc, url, subAddress, tooltip);

        MarkModified(context);

        return BuildResultMessage(text, url, subAddress, tooltip, paragraphIndex);
    }

    private static void ValidateUrlIfPresent(string? url)
    {
        if (!string.IsNullOrEmpty(url))
            WordHyperlinkHelper.ValidateUrlFormat(url);
    }

    private static void PositionBuilder(Document doc, DocumentBuilder builder, int? paragraphIndex)
    {
        if (!paragraphIndex.HasValue)
        {
            builder.MoveToDocumentEnd();
            return;
        }

        if (paragraphIndex.Value == -1)
            PositionAtDocumentStart(doc, builder);
        else
            PositionAfterParagraph(doc, builder, paragraphIndex.Value);
    }

    private static void PositionAtDocumentStart(Document doc, DocumentBuilder builder)
    {
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        if (paragraphs.Count > 0 && paragraphs[0] is WordParagraph firstPara)
        {
            var newPara = new WordParagraph(doc);
            doc.FirstSection.Body.InsertBefore(newPara, firstPara);
            builder.MoveTo(newPara);
        }
        else
        {
            builder.MoveToDocumentStart();
        }
    }

    private static void PositionAfterParagraph(Document doc, DocumentBuilder builder, int paragraphIndex)
    {
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
            throw new ArgumentException(
                $"Paragraph index {paragraphIndex} is out of range (document has {paragraphs.Count} paragraphs)");

        if (paragraphs[paragraphIndex] is not WordParagraph targetPara)
            throw new InvalidOperationException($"Unable to find paragraph at index {paragraphIndex}");

        var parentNode = targetPara.ParentNode;
        if (parentNode == null)
            throw new InvalidOperationException($"Unable to find parent node of paragraph at index {paragraphIndex}");

        var newPara = new WordParagraph(doc);
        parentNode.InsertAfter(newPara, targetPara);
        builder.MoveTo(newPara);
    }

    private static void InsertHyperlink(DocumentBuilder builder, string? text, string? url, string? subAddress)
    {
        if (!string.IsNullOrEmpty(subAddress))
            builder.InsertHyperlink(text ?? "", subAddress, true);
        else
            builder.InsertHyperlink(text ?? "", url!, false);
    }

    private static void ApplyHyperlinkSettings(Document doc, string? url, string? subAddress, string? tooltip)
    {
        var fields = doc.Range.Fields;
        if (fields.Count == 0) return;

        if (fields[^1] is not FieldHyperlink hyperlinkField) return;

        if (!string.IsNullOrEmpty(tooltip))
            hyperlinkField.ScreenTip = tooltip;

        if (!string.IsNullOrEmpty(url) && !string.IsNullOrEmpty(subAddress))
        {
            hyperlinkField.Address = url;
            hyperlinkField.SubAddress = subAddress;
        }
    }

    private static string BuildResultMessage(string? text, string? url, string? subAddress, string? tooltip,
        int? paragraphIndex)
    {
        var result = "Hyperlink added successfully\n";
        result += $"Display text: {text}\n";
        if (!string.IsNullOrEmpty(url)) result += $"URL: {url}\n";
        if (!string.IsNullOrEmpty(subAddress)) result += $"SubAddress (bookmark): {subAddress}\n";
        if (!string.IsNullOrEmpty(tooltip)) result += $"Tooltip: {tooltip}\n";
        result += GetInsertPositionMessage(paragraphIndex);
        return result;
    }

    private static string GetInsertPositionMessage(int? paragraphIndex)
    {
        if (!paragraphIndex.HasValue)
            return "Insert position: end of document";

        return paragraphIndex.Value == -1
            ? "Insert position: beginning of document"
            : $"Insert position: after paragraph #{paragraphIndex.Value}";
    }
}
