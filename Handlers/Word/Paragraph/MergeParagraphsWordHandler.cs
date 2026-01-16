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
        var mergeParams = ExtractMergeParagraphsParameters(parameters);

        if (!mergeParams.StartParagraphIndex.HasValue)
            throw new ArgumentException("startParagraphIndex parameter is required for merge operation");
        if (!mergeParams.EndParagraphIndex.HasValue)
            throw new ArgumentException("endParagraphIndex parameter is required for merge operation");

        var doc = context.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        if (mergeParams.StartParagraphIndex.Value < 0 || mergeParams.StartParagraphIndex.Value >= paragraphs.Count)
            throw new ArgumentException(
                $"Start paragraph index {mergeParams.StartParagraphIndex.Value} is out of range. The document has {paragraphs.Count} paragraphs (valid indices: 0-{paragraphs.Count - 1}).");

        if (mergeParams.EndParagraphIndex.Value < 0 || mergeParams.EndParagraphIndex.Value >= paragraphs.Count)
            throw new ArgumentException(
                $"End paragraph index {mergeParams.EndParagraphIndex.Value} is out of range. The document has {paragraphs.Count} paragraphs (valid indices: 0-{paragraphs.Count - 1}).");

        if (mergeParams.StartParagraphIndex.Value > mergeParams.EndParagraphIndex.Value)
            throw new ArgumentException(
                $"Start paragraph index {mergeParams.StartParagraphIndex.Value} cannot be greater than end paragraph index {mergeParams.EndParagraphIndex.Value}");

        if (mergeParams.StartParagraphIndex.Value == mergeParams.EndParagraphIndex.Value)
            throw new ArgumentException("Start and end paragraph indices are the same, no merge needed");

        var startPara = paragraphs[mergeParams.StartParagraphIndex.Value] as Aspose.Words.Paragraph;
        if (startPara == null) throw new InvalidOperationException("Unable to get start paragraph");

        for (var i = mergeParams.StartParagraphIndex.Value + 1; i <= mergeParams.EndParagraphIndex.Value; i++)
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
        result +=
            $"Merge range: Paragraph #{mergeParams.StartParagraphIndex.Value} to #{mergeParams.EndParagraphIndex.Value}\n";
        result +=
            $"Merged paragraphs: {mergeParams.EndParagraphIndex.Value - mergeParams.StartParagraphIndex.Value + 1}\n";
        result += $"Remaining paragraphs: {doc.GetChildNodes(NodeType.Paragraph, true).Count}";

        return result;
    }

    /// <summary>
    ///     Extracts merge paragraphs parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted merge paragraphs parameters.</returns>
    private static MergeParagraphsParameters ExtractMergeParagraphsParameters(OperationParameters parameters)
    {
        return new MergeParagraphsParameters(
            parameters.GetOptional<int?>("startParagraphIndex"),
            parameters.GetOptional<int?>("endParagraphIndex")
        );
    }

    /// <summary>
    ///     Record to hold merge paragraphs parameters.
    /// </summary>
    /// <param name="StartParagraphIndex">The start paragraph index.</param>
    /// <param name="EndParagraphIndex">The end paragraph index.</param>
    private record MergeParagraphsParameters(int? StartParagraphIndex, int? EndParagraphIndex);
}
