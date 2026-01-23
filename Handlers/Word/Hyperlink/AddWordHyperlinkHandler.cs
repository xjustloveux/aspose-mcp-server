using Aspose.Words;
using Aspose.Words.Fields;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Hyperlink;

/// <summary>
///     Handler for adding hyperlinks to Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractAddHyperlinkParameters(parameters);

        if (string.IsNullOrEmpty(p.Url) && string.IsNullOrEmpty(p.SubAddress))
            throw new ArgumentException("Either 'url' or 'subAddress' must be provided for add operation");

        ValidateUrlIfPresent(p.Url);

        var doc = context.Document;
        var builder = new DocumentBuilder(doc);

        PositionBuilder(doc, builder, p.ParagraphIndex);
        InsertHyperlink(builder, p.Text, p.Url, p.SubAddress);
        ApplyHyperlinkSettings(doc, p.Url, p.SubAddress, p.Tooltip);

        MarkModified(context);

        return BuildResultMessage(p);
    }

    /// <summary>
    ///     Extracts add hyperlink parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted add hyperlink parameters.</returns>
    private static AddHyperlinkParameters ExtractAddHyperlinkParameters(OperationParameters parameters)
    {
        return new AddHyperlinkParameters(
            parameters.GetOptional<string?>("text"),
            parameters.GetOptional<string?>("url"),
            parameters.GetOptional<string?>("subAddress"),
            parameters.GetOptional<int?>("paragraphIndex"),
            parameters.GetOptional<string?>("tooltip")
        );
    }

    /// <summary>
    ///     Validates URL format if a URL is provided.
    /// </summary>
    /// <param name="url">The URL to validate.</param>
    private static void ValidateUrlIfPresent(string? url)
    {
        if (!string.IsNullOrEmpty(url))
            WordHyperlinkHelper.ValidateUrlFormat(url);
    }

    /// <summary>
    ///     Positions the document builder at the correct location.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="builder">The document builder.</param>
    /// <param name="paragraphIndex">The paragraph index to position at.</param>
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

    /// <summary>
    ///     Positions the builder at the start of the document.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="builder">The document builder.</param>
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

    /// <summary>
    ///     Positions the builder after the specified paragraph.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="builder">The document builder.</param>
    /// <param name="paragraphIndex">The paragraph index to position after.</param>
    /// <exception cref="ArgumentException">Thrown when paragraph index is out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when paragraph or parent node cannot be found.</exception>
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

    /// <summary>
    ///     Inserts a hyperlink using the document builder.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="text">The display text.</param>
    /// <param name="url">The URL.</param>
    /// <param name="subAddress">The sub-address (bookmark).</param>
    private static void InsertHyperlink(DocumentBuilder builder, string? text, string? url, string? subAddress)
    {
        if (!string.IsNullOrEmpty(subAddress))
            builder.InsertHyperlink(text ?? "", subAddress, true);
        else
            builder.InsertHyperlink(text ?? "", url!, false);
    }

    /// <summary>
    ///     Applies additional settings to the inserted hyperlink.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="url">The URL.</param>
    /// <param name="subAddress">The sub-address.</param>
    /// <param name="tooltip">The tooltip text.</param>
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

    /// <summary>
    ///     Builds the result message for a successful hyperlink addition.
    /// </summary>
    /// <param name="p">The add hyperlink parameters.</param>
    /// <returns>A formatted result message.</returns>
    private static SuccessResult BuildResultMessage(AddHyperlinkParameters p)
    {
        var message = "Hyperlink added successfully\n";
        message += $"Display text: {p.Text}\n";
        if (!string.IsNullOrEmpty(p.Url)) message += $"URL: {p.Url}\n";
        if (!string.IsNullOrEmpty(p.SubAddress)) message += $"SubAddress (bookmark): {p.SubAddress}\n";
        if (!string.IsNullOrEmpty(p.Tooltip)) message += $"Tooltip: {p.Tooltip}\n";
        message += GetInsertPositionMessage(p.ParagraphIndex);
        return new SuccessResult { Message = message };
    }

    /// <summary>
    ///     Gets the insert position description message.
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
    ///     Record to hold add hyperlink parameters.
    /// </summary>
    private sealed record AddHyperlinkParameters(
        string? Text,
        string? Url,
        string? SubAddress,
        int? ParagraphIndex,
        string? Tooltip);
}
