using System.Text;
using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Text;

/// <summary>
///     Represents the location of text within paragraphs.
/// </summary>
/// <param name="ParagraphIndex">The paragraph index where text was found.</param>
/// <param name="EndParagraphIndex">The end paragraph index.</param>
/// <param name="StartRunIndex">The starting run index.</param>
/// <param name="EndRunIndex">The ending run index.</param>
internal record TextLocation(int? ParagraphIndex, int? EndParagraphIndex, int StartRunIndex, int? EndRunIndex);

/// <summary>
///     Represents the start and end run indices for a text range.
/// </summary>
/// <param name="StartIndex">The starting run index.</param>
/// <param name="EndIndex">The ending run index.</param>
internal record RunRangeIndices(int StartIndex, int EndIndex);

/// <summary>
///     Handler for deleting text from Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
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
            var textLocation = FindTextLocation(paragraphs, p.SearchText);
            if (!textLocation.ParagraphIndex.HasValue)
                throw new ArgumentException(
                    $"Text '{p.SearchText}' not found. Please use search operation to confirm text location first.");

            startParagraphIndex = textLocation.ParagraphIndex;
            endParagraphIndex = textLocation.EndParagraphIndex;
            startRunIndex = textLocation.StartRunIndex;
            endRunIndex = textLocation.EndRunIndex;
        }
        else
        {
            if (!startParagraphIndex.HasValue)
                throw new ArgumentException("startParagraphIndex is required when searchText is not provided");
            if (!endParagraphIndex.HasValue)
                throw new ArgumentException("endParagraphIndex is required when searchText is not provided");

            // Delete is destructive, so reject negative indices outright instead of inheriting the
            // resolver's -1 = last-paragraph convention: a caller that miscomputes an index to -1
            // must not silently delete the final paragraph.
            if (startParagraphIndex.Value < 0 || endParagraphIndex.Value < 0)
                throw new ArgumentException("startParagraphIndex and endParagraphIndex must be non-negative");

            // Resolve the caller's story-relative start/end to nodes, then map to document-order
            // positions so the (global) deletion logic and #7 field guards operate unchanged.
            startParagraphIndex = paragraphs.IndexOf(
                ParagraphResolver.Resolve(doc, ParagraphAddress.From(parameters, startParagraphIndex.Value)).Paragraph);
            endParagraphIndex = paragraphs.IndexOf(
                ParagraphResolver.Resolve(doc, ParagraphAddress.From(parameters, endParagraphIndex.Value)).Paragraph);
        }

        // ReSharper disable once RedundantSuppressNullableWarningExpression - on the searchText path the
        // compiler cannot see that FindTextLocation guarantees EndParagraphIndex alongside ParagraphIndex,
        // so the suppression is required to avoid CS8629
        ValidateIndices(paragraphs, startParagraphIndex.Value, endParagraphIndex!.Value);

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
    /// <param name="paragraphs">The collection of paragraphs to search.</param>
    /// <param name="searchText">The text to search for.</param>
    /// <returns>A TextLocation containing the paragraph and run indices.</returns>
    private static TextLocation FindTextLocation(NodeCollection paragraphs, string searchText)
    {
        for (var p = 0; p < paragraphs.Count; p++)
        {
            if (paragraphs[p] is not WordParagraph para) continue;

            var paraText = para.GetText();
            var textIndex = paraText.IndexOf(searchText, StringComparison.OrdinalIgnoreCase);

            if (textIndex >= 0)
            {
                var runs = para.GetChildNodes(NodeType.Run, false);
                var runRange = FindRunRange(runs, textIndex, searchText.Length);
                return new TextLocation(p, p, runRange.StartIndex, runRange.EndIndex);
            }
        }

        return new TextLocation(null, null, 0, null);
    }

    /// <summary>
    ///     Finds the run range for a text position.
    /// </summary>
    /// <param name="runs">The collection of runs to search.</param>
    /// <param name="textIndex">The starting text index.</param>
    /// <param name="textLength">The length of the text.</param>
    /// <returns>A RunRangeIndices containing the start and end run indices.</returns>
    private static RunRangeIndices FindRunRange(NodeCollection runs, int textIndex, int textLength)
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

        return new RunRangeIndices(startRunIdx, endRunIdx);
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
    /// <returns>The resulting text.</returns>
    private static string ExtractDeletedText(NodeCollection paragraphs, WordParagraph startPara, WordParagraph endPara,
        int startParagraphIndex, int endParagraphIndex, int startRunIndex, int? endRunIndex)
    {
        try
        {
            return startParagraphIndex == endParagraphIndex
                ? ExtractFromSameParagraph(startPara, startRunIndex, endRunIndex)
                : ExtractFromMultipleParagraphs(paragraphs, startPara, endPara, startParagraphIndex, endParagraphIndex,
                    startRunIndex, endRunIndex);
        }
        catch
        {
            return "";
        }
    }

    /// <summary>
    ///     Extracts deleted text when start and end are in the same paragraph.
    /// </summary>
    /// <returns>The resulting text.</returns>
    private static string ExtractFromSameParagraph(WordParagraph para, int startRunIndex, int? endRunIndex)
    {
        var runs = para.GetChildNodes(NodeType.Run, false);
        if (runs is not { Count: > 0 }) return "";

        var actualEndRunIndex = endRunIndex ?? runs.Count - 1;
        if (!IsValidRunRange(startRunIndex, actualEndRunIndex, runs.Count)) return "";

        var sb = new StringBuilder();
        for (var i = startRunIndex; i <= actualEndRunIndex; i++)
            if (runs[i] is Run run)
                sb.Append(run.Text);
        return sb.ToString();
    }

    /// <summary>
    ///     Extracts deleted text spanning multiple paragraphs.
    /// </summary>
    /// <returns>The resulting text.</returns>
    private static string ExtractFromMultipleParagraphs(NodeCollection paragraphs, WordParagraph startPara,
        WordParagraph endPara, int startParagraphIndex, int endParagraphIndex, int startRunIndex, int? endRunIndex)
    {
        var sb = new StringBuilder();

        var startRuns = startPara.GetChildNodes(NodeType.Run, false);
        if (startRuns != null && startRuns.Count > startRunIndex)
            for (var i = startRunIndex; i < startRuns.Count; i++)
                if (startRuns[i] is Run run)
                    sb.Append(run.Text);

        for (var p = startParagraphIndex + 1; p < endParagraphIndex; p++)
            if (paragraphs[p] is WordParagraph para)
                sb.Append(para.GetText());

        var endRuns = endPara.GetChildNodes(NodeType.Run, false);
        if (endRuns is { Count: > 0 })
        {
            var actualEndRunIndex = endRunIndex ?? endRuns.Count - 1;
            for (var i = 0; i <= actualEndRunIndex && i < endRuns.Count; i++)
                if (endRuns[i] is Run run)
                    sb.Append(run.Text);
        }

        return sb.ToString();
    }

    /// <summary>
    ///     Validates if the run range is within bounds.
    /// </summary>
    /// <returns><c>true</c> when it does; otherwise <c>false</c>.</returns>
    private static bool IsValidRunRange(int startRunIndex, int endRunIndex, int runCount)
    {
        return startRunIndex >= 0 && startRunIndex < runCount &&
               endRunIndex >= 0 && endRunIndex < runCount &&
               startRunIndex <= endRunIndex;
    }

    /// <summary>
    ///     Performs the actual text deletion.
    /// </summary>
    private static void DeleteTextRange(NodeCollection paragraphs, WordParagraph startPara, WordParagraph endPara,
        int startParagraphIndex, int endParagraphIndex, int startRunIndex, int? endRunIndex)
    {
        if (startParagraphIndex == endParagraphIndex)
            DeleteFromSameParagraph(startPara, startRunIndex, endRunIndex);
        else
            DeleteFromMultipleParagraphs(paragraphs, startPara, endPara, startParagraphIndex, endParagraphIndex,
                startRunIndex, endRunIndex);
    }

    /// <summary>
    ///     Deletes runs from a single paragraph.
    /// </summary>
    private static void DeleteFromSameParagraph(WordParagraph para, int startRunIndex, int? endRunIndex)
    {
        var runs = para.GetChildNodes(NodeType.Run, false);
        if (runs is not { Count: > 0 }) return;

        var actualEndRunIndex = endRunIndex ?? runs.Count - 1;
        if (!IsValidRunRange(startRunIndex, actualEndRunIndex, runs.Count)) return;

        // Indexed before the loop. Removing runs does not invalidate it: the positions are a
        // snapshot, and every run still asked about is one it already knows.
        var extents = FieldBoundaryHelper.FieldExtents.Of(para.Document as Document
                                                          ?? throw new InvalidOperationException(
                                                              "The paragraph is not part of a document."));

        for (var i = actualEndRunIndex; i >= startRunIndex; i--)
            if (runs[i] is Run run && extents.EnclosingField(run) == null)
                run.Remove();
    }

    /// <summary>
    ///     Deletes content spanning multiple paragraphs.
    /// </summary>
    private static void DeleteFromMultipleParagraphs(NodeCollection paragraphs, WordParagraph startPara,
        WordParagraph endPara, int startParagraphIndex, int endParagraphIndex, int startRunIndex, int? endRunIndex)
    {
        var extents = FieldBoundaryHelper.FieldExtents.Of(startPara.Document as Document
                                                          ?? throw new InvalidOperationException(
                                                              "The paragraph is not part of a document."));

        // The run loops below skip a run inside a field; the paragraph loop removed whole
        // paragraphs with no such check. A field contained in one of them goes with it, which is a
        // complete removal — but a field that runs *through* one loses a marker and leaves the
        // document with a field that has no end. Checked here, before the first run is touched:
        // checking it at the paragraph loop left the start paragraph already truncated when the
        // refusal came, so a refused delete had still changed the document (§21.3).
        for (var p = endParagraphIndex - 1; p > startParagraphIndex; p--)
            if (paragraphs[p] is WordParagraph crossed && extents.WouldSplitAField(crossed))
                throw new ArgumentException(
                    "The range crosses a field that begins or ends outside it, so deleting it "
                    + "would leave the field without one of its markers. Delete the field with "
                    + "word_field, or choose a range that does not cut through one.");

        var startRuns = startPara.GetChildNodes(NodeType.Run, false);
        if (startRuns != null && startRuns.Count > startRunIndex)
            for (var i = startRuns.Count - 1; i >= startRunIndex; i--)
                if (startRuns[i] is Run run && extents.EnclosingField(run) == null)
                    run.Remove();

        for (var p = endParagraphIndex - 1; p > startParagraphIndex; p--)
            paragraphs[p]?.Remove();

        var endRuns = endPara.GetChildNodes(NodeType.Run, false);
        if (endRuns is { Count: > 0 })
        {
            var actualEndRunIndex = endRunIndex ?? endRuns.Count - 1;
            for (var i = actualEndRunIndex; i >= 0; i--)
                if (i < endRuns.Count && endRuns[i] is Run run &&
                    extents.EnclosingField(run) == null)
                    run.Remove();
        }
    }

    /// <summary>
    ///     Builds the result message.
    /// </summary>
    /// <returns>The success result.</returns>
    private static SuccessResult BuildResultMessage(string? searchText, int startPara, int startRun,
        int endPara, int? endRun, string deletedText)
    {
        var preview = deletedText.Length > 50 ? deletedText.Substring(0, 50) + "..." : deletedText;

        var message = "Text deleted successfully.";
        if (!string.IsNullOrEmpty(searchText))
            message += $" Deleted text: {searchText}.";
        message += $" Range: Paragraph {startPara} Run {startRun} to Paragraph {endPara} Run {endRun ?? -1}.";
        if (!string.IsNullOrEmpty(preview))
            message += $" Preview: {preview}";

        return new SuccessResult { Message = message };
    }

    private sealed record DeleteParameters(
        string? SearchText,
        int? StartParagraphIndex,
        int StartRunIndex,
        int? EndParagraphIndex,
        int? EndRunIndex);
}
