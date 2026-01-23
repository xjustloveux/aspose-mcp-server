using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Paragraph;

/// <summary>
///     Handler for copying paragraph format in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var copyParams = ExtractCopyParagraphFormatParameters(parameters);

        if (!copyParams.SourceParagraphIndex.HasValue)
            throw new ArgumentException("sourceParagraphIndex parameter is required for copy_format operation");
        if (!copyParams.TargetParagraphIndex.HasValue)
            throw new ArgumentException("targetParagraphIndex parameter is required for copy_format operation");

        var doc = context.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        if (copyParams.SourceParagraphIndex.Value < 0 || copyParams.SourceParagraphIndex.Value >= paragraphs.Count)
            throw new ArgumentException(
                $"Source paragraph index {copyParams.SourceParagraphIndex.Value} is out of range. The document has {paragraphs.Count} paragraphs (valid indices: 0-{paragraphs.Count - 1}).");

        if (copyParams.TargetParagraphIndex.Value < 0 || copyParams.TargetParagraphIndex.Value >= paragraphs.Count)
            throw new ArgumentException(
                $"Target paragraph index {copyParams.TargetParagraphIndex.Value} is out of range. The document has {paragraphs.Count} paragraphs (valid indices: 0-{paragraphs.Count - 1}).");

        var sourcePara = paragraphs[copyParams.SourceParagraphIndex.Value] as Aspose.Words.Paragraph;
        var targetPara = paragraphs[copyParams.TargetParagraphIndex.Value] as Aspose.Words.Paragraph;

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

        var message = "Paragraph format copied successfully\n";
        message += $"Source paragraph: #{copyParams.SourceParagraphIndex.Value}\n";
        message += $"Target paragraph: #{copyParams.TargetParagraphIndex.Value}";

        return new SuccessResult { Message = message };
    }

    /// <summary>
    ///     Extracts copy paragraph format parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted copy paragraph format parameters.</returns>
    private static CopyParagraphFormatParameters ExtractCopyParagraphFormatParameters(OperationParameters parameters)
    {
        return new CopyParagraphFormatParameters(
            parameters.GetOptional<int?>("sourceParagraphIndex"),
            parameters.GetOptional<int?>("targetParagraphIndex")
        );
    }

    /// <summary>
    ///     Record to hold copy paragraph format parameters.
    /// </summary>
    /// <param name="SourceParagraphIndex">The source paragraph index.</param>
    /// <param name="TargetParagraphIndex">The target paragraph index.</param>
    private sealed record CopyParagraphFormatParameters(int? SourceParagraphIndex, int? TargetParagraphIndex);
}
