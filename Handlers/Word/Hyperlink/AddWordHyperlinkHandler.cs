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

        if (!string.IsNullOrEmpty(url))
            WordHyperlinkHelper.ValidateUrlFormat(url);

        var doc = context.Document;
        var builder = new DocumentBuilder(doc);

        if (paragraphIndex.HasValue)
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            if (paragraphIndex.Value == -1)
            {
                if (paragraphs.Count > 0)
                {
                    if (paragraphs[0] is WordParagraph firstPara)
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
                else
                {
                    builder.MoveToDocumentStart();
                }
            }
            else if (paragraphIndex.Value >= 0 && paragraphIndex.Value < paragraphs.Count)
            {
                if (paragraphs[paragraphIndex.Value] is WordParagraph targetPara)
                {
                    var newPara = new WordParagraph(doc);
                    var parentNode = targetPara.ParentNode;
                    if (parentNode != null)
                    {
                        parentNode.InsertAfter(newPara, targetPara);
                        builder.MoveTo(newPara);
                    }
                    else
                    {
                        throw new InvalidOperationException(
                            $"Unable to find parent node of paragraph at index {paragraphIndex.Value}");
                    }
                }
                else
                {
                    throw new InvalidOperationException(
                        $"Unable to find paragraph at index {paragraphIndex.Value}");
                }
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

        if (!string.IsNullOrEmpty(subAddress))
            builder.InsertHyperlink(text ?? "", subAddress, true);
        else
            builder.InsertHyperlink(text ?? "", url!, false);

        var fields = doc.Range.Fields;
        if (fields.Count > 0)
        {
            var lastField = fields[^1];
            if (lastField is FieldHyperlink hyperlinkField)
            {
                if (!string.IsNullOrEmpty(tooltip))
                    hyperlinkField.ScreenTip = tooltip;
                if (!string.IsNullOrEmpty(url) && !string.IsNullOrEmpty(subAddress))
                {
                    hyperlinkField.Address = url;
                    hyperlinkField.SubAddress = subAddress;
                }
            }
        }

        MarkModified(context);

        var result = "Hyperlink added successfully\n";
        result += $"Display text: {text}\n";
        if (!string.IsNullOrEmpty(url)) result += $"URL: {url}\n";
        if (!string.IsNullOrEmpty(subAddress)) result += $"SubAddress (bookmark): {subAddress}\n";
        if (!string.IsNullOrEmpty(tooltip)) result += $"Tooltip: {tooltip}\n";
        if (paragraphIndex.HasValue)
        {
            if (paragraphIndex.Value == -1)
                result += "Insert position: beginning of document";
            else
                result += $"Insert position: after paragraph #{paragraphIndex.Value}";
        }
        else
        {
            result += "Insert position: end of document";
        }

        return result;
    }
}
