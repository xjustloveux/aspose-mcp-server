using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Paragraph;

/// <summary>
///     Handler for inserting paragraphs in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var insertParams = ExtractInsertParameters(parameters);
        if (string.IsNullOrEmpty(insertParams.Text))
            throw new ArgumentException("text parameter is required for insert operation");

        var doc = context.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        var (targetPara, insertPosition) = FindInsertTarget(paragraphs, insertParams.ParagraphIndex);
        var para = CreateParagraph(doc, insertParams);
        InsertParagraph(doc, para, targetPara, insertParams.ParagraphIndex);

        MarkModified(context);

        return BuildResultMessage(insertPosition, insertParams, paragraphs.Count + 1);
    }

    /// <summary>
    ///     Extracts insert parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted insert parameters.</returns>
    private static InsertParameters ExtractInsertParameters(OperationParameters parameters)
    {
        return new InsertParameters(
            parameters.GetOptional<string?>("text"),
            parameters.GetOptional<int?>("paragraphIndex"),
            parameters.GetOptional<string?>("styleName"),
            parameters.GetOptional<string?>("alignment"),
            parameters.GetOptional<double?>("indentLeft"),
            parameters.GetOptional<double?>("indentRight"),
            parameters.GetOptional<double?>("firstLineIndent"),
            parameters.GetOptional<double?>("spaceBefore"),
            parameters.GetOptional<double?>("spaceAfter")
        );
    }

    /// <summary>
    ///     Finds the target position for insertion.
    /// </summary>
    /// <param name="paragraphs">The collection of paragraphs.</param>
    /// <param name="paragraphIndex">The paragraph index.</param>
    /// <returns>A tuple containing the target paragraph and position description.</returns>
    /// <exception cref="ArgumentException">Thrown when paragraph index is out of range.</exception>
    private static (WordParagraph? targetPara, string insertPosition) FindInsertTarget(NodeCollection paragraphs,
        int? paragraphIndex)
    {
        if (!paragraphIndex.HasValue)
            return (null, "end of document");

        if (paragraphIndex.Value == -1 && paragraphs.Count > 0)
            return (paragraphs[0] as WordParagraph, "beginning of document");

        if (paragraphIndex.Value >= 0 && paragraphIndex.Value < paragraphs.Count)
            return (paragraphs[paragraphIndex.Value] as WordParagraph, $"after paragraph #{paragraphIndex.Value}");

        var validRange = paragraphs.Count > 0 ? $"0-{paragraphs.Count - 1}" : "none (document is empty)";
        throw new ArgumentException(
            $"Paragraph index {paragraphIndex.Value} is out of range. The document has {paragraphs.Count} paragraphs (valid indices: {validRange}, or -1 for beginning).");
    }

    /// <summary>
    ///     Creates a new paragraph with the specified settings.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="p">The insert parameters.</param>
    /// <returns>The created paragraph.</returns>
    private static WordParagraph CreateParagraph(Document doc, InsertParameters p)
    {
        var para = new WordParagraph(doc);
        para.AppendChild(new Run(doc, p.Text));

        ApplyStyle(doc, para, p.StyleName);
        ApplyFormatting(para, p);

        return para;
    }

    /// <summary>
    ///     Applies a style to the paragraph.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="para">The paragraph.</param>
    /// <param name="styleName">The style name.</param>
    /// <exception cref="ArgumentException">Thrown when style is not found.</exception>
    private static void ApplyStyle(Document doc, WordParagraph para, string? styleName)
    {
        if (string.IsNullOrEmpty(styleName)) return;
        var style = doc.Styles[styleName];
        if (style == null)
            throw new ArgumentException(
                $"Style '{styleName}' not found. Use word_get_styles tool to view available styles");
        para.ParagraphFormat.StyleName = styleName;
    }

    /// <summary>
    ///     Applies formatting to the paragraph.
    /// </summary>
    /// <param name="para">The paragraph.</param>
    /// <param name="p">The insert parameters.</param>
    private static void ApplyFormatting(WordParagraph para, InsertParameters p)
    {
        if (!string.IsNullOrEmpty(p.Alignment))
            para.ParagraphFormat.Alignment = WordParagraphHelper.GetAlignment(p.Alignment);
        if (p.IndentLeft.HasValue) para.ParagraphFormat.LeftIndent = p.IndentLeft.Value;
        if (p.IndentRight.HasValue) para.ParagraphFormat.RightIndent = p.IndentRight.Value;
        if (p.FirstLineIndent.HasValue) para.ParagraphFormat.FirstLineIndent = p.FirstLineIndent.Value;
        if (p.SpaceBefore.HasValue) para.ParagraphFormat.SpaceBefore = p.SpaceBefore.Value;
        if (p.SpaceAfter.HasValue) para.ParagraphFormat.SpaceAfter = p.SpaceAfter.Value;
    }

    /// <summary>
    ///     Inserts the paragraph at the target position.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="para">The paragraph to insert.</param>
    /// <param name="targetPara">The target paragraph.</param>
    /// <param name="paragraphIndex">The paragraph index.</param>
    private static void InsertParagraph(Document doc, WordParagraph para, WordParagraph? targetPara,
        int? paragraphIndex)
    {
        if (targetPara == null)
        {
            doc.FirstSection.Body.AppendChild(para);
            return;
        }

        if (paragraphIndex!.Value == -1)
            targetPara.ParentNode.InsertBefore(para, targetPara);
        else
            targetPara.ParentNode.InsertAfter(para, targetPara);
    }

    /// <summary>
    ///     Builds the result message for a successful insertion.
    /// </summary>
    /// <param name="insertPosition">The insert position description.</param>
    /// <param name="p">The insert parameters.</param>
    /// <param name="newParagraphCount">The new paragraph count.</param>
    /// <returns>The result message.</returns>
    private static SuccessResult BuildResultMessage(string insertPosition, InsertParameters p, int newParagraphCount)
    {
        var message = "Paragraph inserted successfully\n";
        message += $"Insert position: {insertPosition}\n";
        if (!string.IsNullOrEmpty(p.StyleName)) message += $"Applied style: {p.StyleName}\n";
        if (!string.IsNullOrEmpty(p.Alignment)) message += $"Alignment: {p.Alignment}\n";
        message += $"Document paragraph count: {newParagraphCount}";
        return new SuccessResult { Message = message };
    }

    /// <summary>
    ///     Record to hold insert parameters.
    /// </summary>
    private sealed record InsertParameters(
        string? Text,
        int? ParagraphIndex,
        string? StyleName,
        string? Alignment,
        double? IndentLeft,
        double? IndentRight,
        double? FirstLineIndent,
        double? SpaceBefore,
        double? SpaceAfter);
}
