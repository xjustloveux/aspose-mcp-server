using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Paragraph;

/// <summary>
///     Handler for merging paragraphs in Word documents.
/// </summary>
public class MergeParagraphsWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "merge";

    /// <summary>
    ///     Merges multiple consecutive paragraphs into one.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: startParagraphIndex, endParagraphIndex
    /// </param>
    /// <returns>Success message with merge details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var startParagraphIndex = parameters.GetOptional<int?>("startParagraphIndex");
        var endParagraphIndex = parameters.GetOptional<int?>("endParagraphIndex");

        if (!startParagraphIndex.HasValue)
            throw new ArgumentException("startParagraphIndex parameter is required for merge operation");
        if (!endParagraphIndex.HasValue)
            throw new ArgumentException("endParagraphIndex parameter is required for merge operation");

        var doc = context.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        if (startParagraphIndex.Value < 0 || startParagraphIndex.Value >= paragraphs.Count)
            throw new ArgumentException(
                $"Start paragraph index {startParagraphIndex.Value} is out of range. The document has {paragraphs.Count} paragraphs (valid indices: 0-{paragraphs.Count - 1}).");

        if (endParagraphIndex.Value < 0 || endParagraphIndex.Value >= paragraphs.Count)
            throw new ArgumentException(
                $"End paragraph index {endParagraphIndex.Value} is out of range. The document has {paragraphs.Count} paragraphs (valid indices: 0-{paragraphs.Count - 1}).");

        if (startParagraphIndex.Value > endParagraphIndex.Value)
            throw new ArgumentException(
                $"Start paragraph index {startParagraphIndex.Value} cannot be greater than end paragraph index {endParagraphIndex.Value}");

        if (startParagraphIndex.Value == endParagraphIndex.Value)
            throw new ArgumentException("Start and end paragraph indices are the same, no merge needed");

        var startPara = paragraphs[startParagraphIndex.Value] as Aspose.Words.Paragraph;
        if (startPara == null) throw new InvalidOperationException("Unable to get start paragraph");

        for (var i = startParagraphIndex.Value + 1; i <= endParagraphIndex.Value; i++)
            if (paragraphs[i] is Aspose.Words.Paragraph para)
            {
                if (startPara.Runs.Count > 0)
                {
                    var spaceRun = new Run(doc, " ");
                    startPara.AppendChild(spaceRun);
                }

                var runsToMove = para.Runs.ToArray();
                foreach (var run in runsToMove) startPara.AppendChild(run);

                para.Remove();
            }

        MarkModified(context);

        var result = "Paragraphs merged successfully\n";
        result += $"Merge range: Paragraph #{startParagraphIndex.Value} to #{endParagraphIndex.Value}\n";
        result += $"Merged paragraphs: {endParagraphIndex.Value - startParagraphIndex.Value + 1}\n";
        result += $"Remaining paragraphs: {doc.GetChildNodes(NodeType.Paragraph, true).Count}";

        return result;
    }
}
