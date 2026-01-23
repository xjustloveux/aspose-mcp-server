using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;
using static Aspose.Words.ConvertUtil;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.List;

/// <summary>
///     Handler for setting list format in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractSetListFormatParameters(parameters);

        var doc = context.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        if (p.ParagraphIndex < 0 || p.ParagraphIndex >= paragraphs.Count)
            throw new ArgumentException(
                $"Paragraph index {p.ParagraphIndex} is out of range (document has {paragraphs.Count} paragraphs)");

        var para = paragraphs[p.ParagraphIndex] as WordParagraph;
        if (para == null)
            throw new InvalidOperationException($"Unable to find paragraph at index {p.ParagraphIndex}");

        List<string> changes = [];

        if (!string.IsNullOrEmpty(p.NumberStyle) && para.ListFormat.IsListItem)
        {
            var list = para.ListFormat.List;
            if (list != null)
            {
                var level = para.ListFormat.ListLevelNumber;
                var listLevel = list.ListLevels[level];

                var style = p.NumberStyle.ToLower() switch
                {
                    "arabic" => NumberStyle.Arabic,
                    "roman" => NumberStyle.UppercaseRoman,
                    "letter" => NumberStyle.UppercaseLetter,
                    "bullet" => NumberStyle.Bullet,
                    "none" => NumberStyle.None,
                    _ => NumberStyle.Arabic
                };

                listLevel.NumberStyle = style;
                changes.Add($"Number style: {p.NumberStyle} (affects all items at level {level} in this list)");
            }
        }

        if (p.IndentLevel.HasValue)
        {
            para.ParagraphFormat.LeftIndent = InchToPoint(0.5 * p.IndentLevel.Value);
            changes.Add($"Indent level: {p.IndentLevel.Value} ({InchToPoint(0.5 * p.IndentLevel.Value):F1} points)");
        }

        if (p.LeftIndent.HasValue)
        {
            para.ParagraphFormat.LeftIndent = p.LeftIndent.Value;
            changes.Add($"Left indent: {p.LeftIndent.Value} points");
        }

        if (p.FirstLineIndent.HasValue)
        {
            para.ParagraphFormat.FirstLineIndent = p.FirstLineIndent.Value;
            changes.Add($"First line indent: {p.FirstLineIndent.Value} points");
        }

        MarkModified(context);

        var result = "List format set successfully\n";
        result += $"Paragraph index: {p.ParagraphIndex}\n";
        if (changes.Count > 0)
            result += $"Changes: {string.Join(", ", changes)}";
        else
            result += "No change parameters provided";

        return new SuccessResult { Message = result };
    }

    private static SetListFormatParameters ExtractSetListFormatParameters(OperationParameters parameters)
    {
        return new SetListFormatParameters(
            parameters.GetRequired<int>("paragraphIndex"),
            parameters.GetOptional<string?>("numberStyle"),
            parameters.GetOptional<int?>("indentLevel"),
            parameters.GetOptional<double?>("leftIndent"),
            parameters.GetOptional<double?>("firstLineIndent"));
    }

    private sealed record SetListFormatParameters(
        int ParagraphIndex,
        string? NumberStyle,
        int? IndentLevel,
        double? LeftIndent,
        double? FirstLineIndent);
}
