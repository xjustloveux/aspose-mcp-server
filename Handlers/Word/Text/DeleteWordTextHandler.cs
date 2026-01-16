using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Text;

/// <summary>
///     Handler for deleting text from Word documents.
/// </summary>
public class DeleteWordTextHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes text from the document by search text or paragraph/run range.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: searchText OR (startParagraphIndex AND endParagraphIndex).
    ///     Optional: startRunIndex, endRunIndex.
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or indices are invalid.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractDeleteParameters(parameters);

        var doc = context.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        var startParagraphIndex = p.StartParagraphIndex;
        var endParagraphIndex = p.EndParagraphIndex;
        var startRunIndex = p.StartRunIndex;
        var endRunIndex = p.EndRunIndex;

        if (!string.IsNullOrEmpty(p.SearchText))
        {
            var (foundStart, foundEnd, foundStartRun, foundEndRun) = FindTextLocation(paragraphs, p.SearchText);
            if (!foundStart.HasValue)
                throw new ArgumentException(
                    $"Text '{p.SearchText}' not found. Please use search operation to confirm text location first.");

            startParagraphIndex = foundStart;
            endParagraphIndex = foundEnd;
            startRunIndex = foundStartRun;
            endRunIndex = foundEndRun;
        }
        else
        {
            if (!startParagraphIndex.HasValue)
                throw new ArgumentException("startParagraphIndex is required when searchText is not provided");
            if (!endParagraphIndex.HasValue)
                throw new ArgumentException("endParagraphIndex is required when searchText is not provided");
        }

        // ReSharper disable once RedundantSuppressNullableWarningExpression - Complex control flow guarantees non-null
        ValidateIndices(paragraphs, startParagraphIndex!.Value, endParagraphIndex!.Value);

        var startPara = (WordParagraph)paragraphs[startParagraphIndex.Value];
        var endPara = (WordParagraph)paragraphs[endParagraphIndex.Value];

        var deletedText = ExtractDeletedText(paragraphs, startPara, endPara, startParagraphIndex.Value,
            endParagraphIndex.Value, startRunIndex, endRunIndex);

        DeleteTextRange(paragraphs, startPara, endPara, startParagraphIndex.Value, endParagraphIndex.Value,
            startRunIndex, endRunIndex);

        MarkModified(context);

        return BuildResultMessage(p.SearchText, startParagraphIndex.Value, startRunIndex,
            endParagraphIndex.Value, endRunIndex, deletedText);
    }

    private static DeleteParameters ExtractDeleteParameters(OperationParameters parameters)
    {
        return new DeleteParameters(
            parameters.GetOptional<string?>("searchText"),
            parameters.GetOptional<int?>("startParagraphIndex"),
            parameters.GetOptional("startRunIndex", 0),
            parameters.GetOptional<int?>("endParagraphIndex"),
            parameters.GetOptional<int?>("endRunIndex"));
    }

    /// <summary>
    ///     Finds the location of search text in paragraphs.
    /// </summary>
    private static (int?, int?, int, int?) FindTextLocation(NodeCollection paragraphs, string searchText)
    {
        for (var p = 0; p < paragraphs.Count; p++)
        {
            if (paragraphs[p] is not WordParagraph para) continue;

            var paraText = para.GetText();
            var textIndex = paraText.IndexOf(searchText, StringComparison.OrdinalIgnoreCase);

            if (textIndex >= 0)
            {
                var runs = para.GetChildNodes(NodeType.Run, false);
                var (startRunIdx, endRunIdx) = FindRunRange(runs, textIndex, searchText.Length);
                return (p, p, startRunIdx, endRunIdx);
            }
        }

        return (null, null, 0, null);
    }

    /// <summary>
    ///     Finds the run range for a text position.
    /// </summary>
    private static (int, int) FindRunRange(NodeCollection runs, int textIndex, int textLength)
    {
        var charCount = 0;
        var startRunIdx = 0;
        var endRunIdx = runs.Count - 1;

        for (var r = 0; r < runs.Count; r++)
        {
            if (runs[r] is not Run run) continue;

            var runLength = run.Text.Length;
            if (charCount + runLength > textIndex)
            {
                startRunIdx = r;
                break;
            }

            charCount += runLength;
        }

        charCount = 0;
        var endTextIndex = textIndex + textLength;
        for (var r = 0; r < runs.Count; r++)
        {
            if (runs[r] is not Run run) continue;

            var runLength = run.Text.Length;
            if (charCount + runLength >= endTextIndex)
            {
                endRunIdx = r;
                break;
            }

            charCount += runLength;
        }

        return (startRunIdx, endRunIdx);
    }

    /// <summary>
    ///     Validates paragraph indices.
    /// </summary>
    private static void ValidateIndices(NodeCollection paragraphs, int startIdx, int endIdx)
    {
        if (startIdx < 0 || startIdx >= paragraphs.Count || endIdx < 0 || endIdx >= paragraphs.Count ||
            startIdx > endIdx)
            throw new ArgumentException(
                $"Paragraph index is out of range (document has {paragraphs.Count} paragraphs)");
    }

    /// <summary>
    ///     Extracts the text that will be deleted for preview.
    /// </summary>
    private static string ExtractDeletedText(NodeCollection paragraphs, WordParagraph startPara, WordParagraph endPara,
        int startParagraphIndex, int endParagraphIndex, int startRunIndex, int? endRunIndex)
    {
        var deletedText = "";
        try
        {
            var startRuns = startPara.GetChildNodes(NodeType.Run, false);
            var endRuns = endPara.GetChildNodes(NodeType.Run, false);

            if (startParagraphIndex == endParagraphIndex)
            {
                if (startRuns is { Count: > 0 })
                {
                    var actualEndRunIndex = endRunIndex ?? startRuns.Count - 1;
                    if (startRunIndex >= 0 && startRunIndex < startRuns.Count &&
                        actualEndRunIndex >= 0 && actualEndRunIndex < startRuns.Count &&
                        startRunIndex <= actualEndRunIndex)
                        for (var i = startRunIndex; i <= actualEndRunIndex; i++)
                            if (startRuns[i] is Run run)
                                deletedText += run.Text;
                }
            }
            else
            {
                if (startRuns != null && startRuns.Count > startRunIndex)
                    for (var i = startRunIndex; i < startRuns.Count; i++)
                        if (startRuns[i] is Run run)
                            deletedText += run.Text;

                for (var p = startParagraphIndex + 1; p < endParagraphIndex; p++)
                    if (paragraphs[p] is WordParagraph para)
                        deletedText += para.GetText();

                if (endRuns is { Count: > 0 })
                {
                    var actualEndRunIndex = endRunIndex ?? endRuns.Count - 1;
                    for (var i = 0; i <= actualEndRunIndex && i < endRuns.Count; i++)
                        if (endRuns[i] is Run run)
                            deletedText += run.Text;
                }
            }
        }
        catch
        {
            // Ignore exceptions when extracting deleted text
        }

        return deletedText;
    }

    /// <summary>
    ///     Performs the actual text deletion.
    /// </summary>
    private static void DeleteTextRange(NodeCollection paragraphs, WordParagraph startPara, WordParagraph endPara,
        int startParagraphIndex, int endParagraphIndex, int startRunIndex, int? endRunIndex)
    {
        if (startParagraphIndex == endParagraphIndex)
        {
            var runs = startPara.GetChildNodes(NodeType.Run, false);
            if (runs is { Count: > 0 })
            {
                var actualEndRunIndex = endRunIndex ?? runs.Count - 1;
                if (startRunIndex >= 0 && startRunIndex < runs.Count &&
                    actualEndRunIndex >= 0 && actualEndRunIndex < runs.Count &&
                    startRunIndex <= actualEndRunIndex)
                    for (var i = actualEndRunIndex; i >= startRunIndex; i--)
                        runs[i]?.Remove();
            }
        }
        else
        {
            var startRuns = startPara.GetChildNodes(NodeType.Run, false);
            if (startRuns != null && startRuns.Count > startRunIndex)
                for (var i = startRuns.Count - 1; i >= startRunIndex; i--)
                    startRuns[i]?.Remove();

            for (var p = endParagraphIndex - 1; p > startParagraphIndex; p--)
                paragraphs[p]?.Remove();

            var endRuns = endPara.GetChildNodes(NodeType.Run, false);
            if (endRuns is { Count: > 0 })
            {
                var actualEndRunIndex = endRunIndex ?? endRuns.Count - 1;
                for (var i = actualEndRunIndex; i >= 0; i--)
                    if (i < endRuns.Count)
                        endRuns[i]?.Remove();
            }
        }
    }

    /// <summary>
    ///     Builds the result message.
    /// </summary>
    private static string BuildResultMessage(string? searchText, int startPara, int startRun,
        int endPara, int? endRun, string deletedText)
    {
        var preview = deletedText.Length > 50 ? deletedText.Substring(0, 50) + "..." : deletedText;

        var result = "Text deleted successfully.";
        if (!string.IsNullOrEmpty(searchText))
            result += $" Deleted text: {searchText}.";
        result += $" Range: Paragraph {startPara} Run {startRun} to Paragraph {endPara} Run {endRun ?? -1}.";
        if (!string.IsNullOrEmpty(preview))
            result += $" Preview: {preview}";

        return Success(result);
    }

    private record DeleteParameters(
        string? SearchText,
        int? StartParagraphIndex,
        int StartRunIndex,
        int? EndParagraphIndex,
        int? EndRunIndex);
}
