using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using static Aspose.Words.ConvertUtil;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.List;

/// <summary>
///     Handler for setting list format in Word documents.
/// </summary>
public class SetWordListFormatHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "set_format";

    /// <summary>
    ///     Sets list formatting options for a paragraph.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: paragraphIndex.
    ///     Optional: numberStyle, indentLevel, leftIndent, firstLineIndent
    /// </param>
    /// <returns>Success message with format changes.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var paragraphIndex = parameters.GetRequired<int>("paragraphIndex");
        var numberStyle = parameters.GetOptional<string?>("numberStyle");
        var indentLevel = parameters.GetOptional<int?>("indentLevel");
        var leftIndent = parameters.GetOptional<double?>("leftIndent");
        var firstLineIndent = parameters.GetOptional<double?>("firstLineIndent");

        var doc = context.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
            throw new ArgumentException(
                $"Paragraph index {paragraphIndex} is out of range (document has {paragraphs.Count} paragraphs)");

        var para = paragraphs[paragraphIndex] as WordParagraph;
        if (para == null)
            throw new InvalidOperationException($"Unable to find paragraph at index {paragraphIndex}");

        List<string> changes = [];

        if (!string.IsNullOrEmpty(numberStyle) && para.ListFormat.IsListItem)
        {
            var list = para.ListFormat.List;
            if (list != null)
            {
                var level = para.ListFormat.ListLevelNumber;
                var listLevel = list.ListLevels[level];

                var style = numberStyle.ToLower() switch
                {
                    "arabic" => NumberStyle.Arabic,
                    "roman" => NumberStyle.UppercaseRoman,
                    "letter" => NumberStyle.UppercaseLetter,
                    "bullet" => NumberStyle.Bullet,
                    "none" => NumberStyle.None,
                    _ => NumberStyle.Arabic
                };

                listLevel.NumberStyle = style;
                changes.Add($"Number style: {numberStyle} (affects all items at level {level} in this list)");
            }
        }

        if (indentLevel.HasValue)
        {
            para.ParagraphFormat.LeftIndent = InchToPoint(0.5 * indentLevel.Value);
            changes.Add($"Indent level: {indentLevel.Value} ({InchToPoint(0.5 * indentLevel.Value):F1} points)");
        }

        if (leftIndent.HasValue)
        {
            para.ParagraphFormat.LeftIndent = leftIndent.Value;
            changes.Add($"Left indent: {leftIndent.Value} points");
        }

        if (firstLineIndent.HasValue)
        {
            para.ParagraphFormat.FirstLineIndent = firstLineIndent.Value;
            changes.Add($"First line indent: {firstLineIndent.Value} points");
        }

        MarkModified(context);

        var result = "List format set successfully\n";
        result += $"Paragraph index: {paragraphIndex}\n";
        if (changes.Count > 0)
            result += $"Changes: {string.Join(", ", changes)}";
        else
            result += "No change parameters provided";

        return Success(result);
    }
}
