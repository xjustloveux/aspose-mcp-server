using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Paragraph;

/// <summary>
///     Handler for copying paragraph format in Word documents.
/// </summary>
public class CopyParagraphFormatWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "copy_format";

    /// <summary>
    ///     Copies formatting from one paragraph to another.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: sourceParagraphIndex, targetParagraphIndex
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var sourceParagraphIndex = parameters.GetOptional<int?>("sourceParagraphIndex");
        var targetParagraphIndex = parameters.GetOptional<int?>("targetParagraphIndex");

        if (!sourceParagraphIndex.HasValue)
            throw new ArgumentException("sourceParagraphIndex parameter is required for copy_format operation");
        if (!targetParagraphIndex.HasValue)
            throw new ArgumentException("targetParagraphIndex parameter is required for copy_format operation");

        var doc = context.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        if (sourceParagraphIndex.Value < 0 || sourceParagraphIndex.Value >= paragraphs.Count)
            throw new ArgumentException(
                $"Source paragraph index {sourceParagraphIndex.Value} is out of range. The document has {paragraphs.Count} paragraphs (valid indices: 0-{paragraphs.Count - 1}).");

        if (targetParagraphIndex.Value < 0 || targetParagraphIndex.Value >= paragraphs.Count)
            throw new ArgumentException(
                $"Target paragraph index {targetParagraphIndex.Value} is out of range. The document has {paragraphs.Count} paragraphs (valid indices: 0-{paragraphs.Count - 1}).");

        var sourcePara = paragraphs[sourceParagraphIndex.Value] as Aspose.Words.Paragraph;
        var targetPara = paragraphs[targetParagraphIndex.Value] as Aspose.Words.Paragraph;

        if (sourcePara == null || targetPara == null)
            throw new InvalidOperationException("Unable to get paragraphs");

        targetPara.ParagraphFormat.StyleName = sourcePara.ParagraphFormat.StyleName;
        targetPara.ParagraphFormat.Alignment = sourcePara.ParagraphFormat.Alignment;
        targetPara.ParagraphFormat.LeftIndent = sourcePara.ParagraphFormat.LeftIndent;
        targetPara.ParagraphFormat.RightIndent = sourcePara.ParagraphFormat.RightIndent;
        targetPara.ParagraphFormat.FirstLineIndent = sourcePara.ParagraphFormat.FirstLineIndent;
        targetPara.ParagraphFormat.SpaceBefore = sourcePara.ParagraphFormat.SpaceBefore;
        targetPara.ParagraphFormat.SpaceAfter = sourcePara.ParagraphFormat.SpaceAfter;
        targetPara.ParagraphFormat.LineSpacing = sourcePara.ParagraphFormat.LineSpacing;
        targetPara.ParagraphFormat.LineSpacingRule = sourcePara.ParagraphFormat.LineSpacingRule;

        targetPara.ParagraphFormat.TabStops.Clear();
        for (var i = 0; i < sourcePara.ParagraphFormat.TabStops.Count; i++)
        {
            var tabStop = sourcePara.ParagraphFormat.TabStops[i];
            targetPara.ParagraphFormat.TabStops.Add(tabStop.Position, tabStop.Alignment, tabStop.Leader);
        }

        MarkModified(context);

        var result = "Paragraph format copied successfully\n";
        result += $"Source paragraph: #{sourceParagraphIndex.Value}\n";
        result += $"Target paragraph: #{targetParagraphIndex.Value}";

        return result;
    }
}
