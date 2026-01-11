using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Paragraph;

/// <summary>
///     Handler for inserting paragraphs in Word documents.
/// </summary>
public class InsertParagraphWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "insert";

    /// <summary>
    ///     Inserts a new paragraph at the specified position.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: text
    ///     Optional: paragraphIndex, styleName, alignment, indentLeft, indentRight, firstLineIndent, spaceBefore, spaceAfter
    /// </param>
    /// <returns>Success message with insertion details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var text = parameters.GetOptional<string?>("text");
        var paragraphIndex = parameters.GetOptional<int?>("paragraphIndex");
        var styleName = parameters.GetOptional<string?>("styleName");
        var alignment = parameters.GetOptional<string?>("alignment");
        var indentLeft = parameters.GetOptional<double?>("indentLeft");
        var indentRight = parameters.GetOptional<double?>("indentRight");
        var firstLineIndent = parameters.GetOptional<double?>("firstLineIndent");
        var spaceBefore = parameters.GetOptional<double?>("spaceBefore");
        var spaceAfter = parameters.GetOptional<double?>("spaceAfter");

        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("text parameter is required for insert operation");

        var doc = context.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        WordParagraph? targetPara = null;
        var insertPosition = "end of document";

        if (paragraphIndex.HasValue)
        {
            if (paragraphIndex.Value == -1)
            {
                if (paragraphs.Count > 0)
                {
                    targetPara = paragraphs[0] as WordParagraph;
                    insertPosition = "beginning of document";
                }
            }
            else if (paragraphIndex.Value >= 0 && paragraphIndex.Value < paragraphs.Count)
            {
                targetPara = paragraphs[paragraphIndex.Value] as WordParagraph;
                insertPosition = $"after paragraph #{paragraphIndex.Value}";
            }
            else
            {
                var validRange = paragraphs.Count > 0 ? $"0-{paragraphs.Count - 1}" : "none (document is empty)";
                throw new ArgumentException(
                    $"Paragraph index {paragraphIndex.Value} is out of range. The document has {paragraphs.Count} paragraphs (valid indices: {validRange}, or -1 for beginning).");
            }
        }

        var para = new WordParagraph(doc);
        var run = new Run(doc, text);
        para.AppendChild(run);

        if (!string.IsNullOrEmpty(styleName))
        {
            var style = doc.Styles[styleName];
            if (style != null)
                para.ParagraphFormat.StyleName = styleName;
            else
                throw new ArgumentException(
                    $"Style '{styleName}' not found. Use word_get_styles tool to view available styles");
        }

        if (!string.IsNullOrEmpty(alignment))
            para.ParagraphFormat.Alignment = WordParagraphHelper.GetAlignment(alignment);

        if (indentLeft.HasValue) para.ParagraphFormat.LeftIndent = indentLeft.Value;
        if (indentRight.HasValue) para.ParagraphFormat.RightIndent = indentRight.Value;
        if (firstLineIndent.HasValue) para.ParagraphFormat.FirstLineIndent = firstLineIndent.Value;
        if (spaceBefore.HasValue) para.ParagraphFormat.SpaceBefore = spaceBefore.Value;
        if (spaceAfter.HasValue) para.ParagraphFormat.SpaceAfter = spaceAfter.Value;

        if (targetPara != null)
        {
            if (paragraphIndex!.Value == -1)
                targetPara.ParentNode.InsertBefore(para, targetPara);
            else
                targetPara.ParentNode.InsertAfter(para, targetPara);
        }
        else
        {
            var body = doc.FirstSection.Body;
            body.AppendChild(para);
        }

        MarkModified(context);

        var result = "Paragraph inserted successfully\n";
        result += $"Insert position: {insertPosition}\n";
        if (!string.IsNullOrEmpty(styleName)) result += $"Applied style: {styleName}\n";
        if (!string.IsNullOrEmpty(alignment)) result += $"Alignment: {alignment}\n";
        result += $"Document paragraph count: {paragraphs.Count + 1}";

        return result;
    }
}
