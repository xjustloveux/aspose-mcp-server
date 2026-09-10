using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Text;

/// <summary>
///     Handler for deleting text range in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class DeleteRangeWordTextHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete_range";

    /// <summary>
    ///     Deletes text within a specified character range.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: startParagraphIndex, startCharIndex, endParagraphIndex, endCharIndex.
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when indices are out of range.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractDeleteRangeParameters(parameters);

        var doc = context.Document;

        var paragraphs =
            ParagraphResolver.GetStoryParagraphs(doc, ParagraphAddress.From(parameters, p.StartParagraphIndex));

        ValidateParagraphIndices(paragraphs, p.StartParagraphIndex, p.EndParagraphIndex);

        var startPara = paragraphs[p.StartParagraphIndex];
        var endPara = paragraphs[p.EndParagraphIndex];

        ValidateRangeOrdering(p, startPara, endPara);

        if (p.StartParagraphIndex == p.EndParagraphIndex)
            DeleteWithinSameParagraph(startPara, p.StartCharIndex, p.EndCharIndex);
        else
            DeleteAcrossParagraphs(paragraphs, startPara, endPara, p.StartParagraphIndex, p.EndParagraphIndex,
                p.StartCharIndex, p.EndCharIndex);

        MarkModified(context);

        return new SuccessResult { Message = "Text range deleted." };
    }

    private static DeleteRangeParameters ExtractDeleteRangeParameters(OperationParameters parameters)
    {
        return new DeleteRangeParameters(
            parameters.GetRequired<int>("startParagraphIndex"),
            parameters.GetRequired<int>("startCharIndex"),
            parameters.GetRequired<int>("endParagraphIndex"),
            parameters.GetRequired<int>("endCharIndex"));
    }

    /// <summary>
    ///     Validates paragraph indices are within range.
    /// </summary>
    /// <param name="paragraphs">The list of paragraphs.</param>
    /// <param name="startIndex">The start paragraph index.</param>
    /// <param name="endIndex">The end paragraph index.</param>
    /// <exception cref="ArgumentException">Thrown when indices are out of range.</exception>
    private static void ValidateParagraphIndices(List<WordParagraph> paragraphs, int startIndex, int endIndex)
    {
        if (startIndex < 0 || startIndex >= paragraphs.Count ||
            endIndex < 0 || endIndex >= paragraphs.Count)
            throw new ArgumentException("Paragraph indices out of range");
    }

    /// <summary>
    ///     Rejects a range that does not describe a forward span inside its paragraphs.
    ///     <para>
    ///         Each index was bounds-checked on its own, so a reversed pair reached the deletion
    ///         logic: across paragraphs it worked from the later one back, and within a paragraph
    ///         the span helper returned early, deleting nothing while the call still reported
    ///         success (R2-C04). Character indices were never checked against the paragraph's own
    ///         length either.
    ///     </para>
    /// </summary>
    /// <param name="p">The requested range.</param>
    /// <param name="startPara">Paragraph the range starts in.</param>
    /// <param name="endPara">Paragraph the range ends in.</param>
    /// <exception cref="ArgumentException">
    ///     Thrown when the range runs backwards or a character index lies outside its paragraph.
    /// </exception>
    private static void ValidateRangeOrdering(DeleteRangeParameters p, WordParagraph startPara,
        WordParagraph endPara)
    {
        if (p.EndParagraphIndex < p.StartParagraphIndex)
            throw new ArgumentException(
                $"endParagraphIndex ({p.EndParagraphIndex}) must not precede startParagraphIndex "
                + $"({p.StartParagraphIndex}).");

        EnsureCharIndexWithin(p.StartCharIndex, startPara, nameof(p.StartCharIndex));
        EnsureCharIndexWithin(p.EndCharIndex, endPara, nameof(p.EndCharIndex));

        if (p.StartParagraphIndex == p.EndParagraphIndex && p.EndCharIndex < p.StartCharIndex)
            throw new ArgumentException(
                $"endCharIndex ({p.EndCharIndex}) must not precede startCharIndex "
                + $"({p.StartCharIndex}) within the same paragraph.");
    }

    /// <summary>
    ///     Rejects a character index that does not address a position in the paragraph.
    /// </summary>
    /// <param name="value">The requested character index.</param>
    /// <param name="para">Paragraph the index applies to.</param>
    /// <param name="name">Parameter name for the error message.</param>
    /// <exception cref="ArgumentException">Thrown when the index is outside the paragraph.</exception>
    private static void EnsureCharIndexWithin(int value, WordParagraph para, string name)
    {
        // The end index is exclusive, so addressing one past the last character is legitimate.
        // Measured over the paragraph's own runs, which is the character space DeleteSpan edits.
        // Paragraph.GetText() also counts the paragraph terminator and any text inside an inline
        // shape, so it accepted indices the deletion could never reach and reported a length the
        // caller could not address (R3-C03).
        var length = WordRunHelper.GetDirectRunTextLength(para);
        if (value < 0 || value > length)
            throw new ArgumentException(
                $"{name} ({value}) is outside the paragraph, which holds {length} character(s).");
    }

    /// <summary>
    ///     Deletes text within a single paragraph.
    /// </summary>
    /// <param name="para">The paragraph.</param>
    /// <param name="startCharIndex">The start character index.</param>
    /// <param name="endCharIndex">The end character index.</param>
    private static void DeleteWithinSameParagraph(WordParagraph para, int startCharIndex, int endCharIndex)
    {
        DeleteSpan(para, startCharIndex, endCharIndex);
    }

    /// <summary>
    ///     Deletes the half-open character span [<paramref name="startCharIndex" />,
    ///     <paramref name="endCharIndex" />) from a paragraph's own runs.
    ///     The span is measured across every direct run so it can cover any number of them, which
    ///     the previous per-run truncation could not: it edited a single run at each end and left
    ///     the rest of the requested range in place while still reporting success. Runs inside a
    ///     field are skipped so field content stays intact, but they still occupy their positions
    ///     in the character space, keeping indices stable for the caller.
    /// </summary>
    /// <param name="para">The paragraph to edit.</param>
    /// <param name="startCharIndex">First character to delete, inclusive.</param>
    /// <param name="endCharIndex">First character to keep after the deletion, exclusive.</param>
    private static void DeleteSpan(WordParagraph para, int startCharIndex, int endCharIndex)
    {
        if (endCharIndex <= startCharIndex) return;

        // Indexed once for the whole span; asking per run would number the document per run.
        var extents = FieldBoundaryHelper.FieldExtents.Of(para.Document as Document
                                                          ?? throw new InvalidOperationException(
                                                              "The paragraph is not part of a document."));

        var offset = 0;
        foreach (var run in WordRunHelper.GetDirectRuns(para))
        {
            var runStart = offset;
            var runEnd = offset + run.Text.Length;
            offset = runEnd;

            if (extents.EnclosingField(run) != null) continue;
            if (runEnd <= startCharIndex || runStart >= endCharIndex) continue;

            var from = Math.Max(0, startCharIndex - runStart);
            var to = Math.Min(run.Text.Length, endCharIndex - runStart);
            var kept = run.Text[..from] + run.Text[to..];

            if (kept.Length == 0)
                run.Remove();
            else
                run.Text = kept;
        }
    }

    /// <summary>
    ///     Deletes text across multiple paragraphs.
    /// </summary>
    /// <param name="paragraphs">The list of paragraphs.</param>
    /// <param name="startPara">The start paragraph.</param>
    /// <param name="endPara">The end paragraph.</param>
    /// <param name="startParagraphIndex">The start paragraph index.</param>
    /// <param name="endParagraphIndex">The end paragraph index.</param>
    /// <param name="startCharIndex">The start character index.</param>
    /// <param name="endCharIndex">The end character index.</param>
    private static void DeleteAcrossParagraphs(List<WordParagraph> paragraphs, WordParagraph startPara,
        WordParagraph endPara,
        int startParagraphIndex, int endParagraphIndex, int startCharIndex, int endCharIndex)
    {
        // Before anything is truncated. Refusing partway through would leave the start paragraph
        // already shortened by an operation that reported failure (§21.3).
        RefuseIfTheRangeCutsAField(paragraphs, startParagraphIndex, endParagraphIndex);

        TruncateStartParagraph(startPara, startCharIndex);
        RemoveMiddleParagraphs(paragraphs, startParagraphIndex, endParagraphIndex);
        TruncateEndParagraph(endPara, endCharIndex);
    }

    /// <summary>
    ///     Truncates text in the start paragraph from the specified position.
    /// </summary>
    /// <param name="para">The paragraph.</param>
    /// <param name="startCharIndex">The character index from which to truncate.</param>
    private static void TruncateStartParagraph(WordParagraph para, int startCharIndex)
    {
        // Everything from startCharIndex to the end of the paragraph is inside the range, which can
        // span any number of runs. The previous version compared a paragraph-relative index against
        // one run's length and edited only the last run, so a multi-run paragraph kept text the
        // caller had asked to delete while the operation still reported success.
        DeleteSpan(para, startCharIndex, WordRunHelper.GetDirectRunTextLength(para));
    }

    /// <summary>
    ///     Removes paragraphs between start and end paragraphs.
    /// </summary>
    /// <param name="paragraphs">The list of paragraphs.</param>
    /// <param name="startIndex">The start paragraph index.</param>
    /// <param name="endIndex">The end paragraph index.</param>
    private static void RemoveMiddleParagraphs(List<WordParagraph> paragraphs, int startIndex, int endIndex)
    {
        for (var i = startIndex + 1; i < endIndex; i++)
            paragraphs[i].Remove();
    }

    /// <summary>
    ///     Refuses a range that would take one marker of a field and leave the other behind.
    /// </summary>
    /// <param name="paragraphs">The list of paragraphs.</param>
    /// <param name="startIndex">The start paragraph index.</param>
    /// <param name="endIndex">The end paragraph index.</param>
    /// <exception cref="ArgumentException">
    ///     Thrown when a paragraph in the range carries part of a field that begins or ends
    ///     outside it. A field contained wholly in a removed paragraph goes with it, which is a
    ///     complete removal; one that runs through the range is cut in half (§21.3).
    /// </exception>
    private static void RefuseIfTheRangeCutsAField(List<WordParagraph> paragraphs,
        int startIndex, int endIndex)
    {
        if (endIndex - startIndex <= 1) return;

        var document = paragraphs[startIndex].Document as Document
                       ?? throw new InvalidOperationException("The paragraph is not part of a document.");
        var extents = FieldBoundaryHelper.FieldExtents.Of(document);

        for (var i = startIndex + 1; i < endIndex; i++)
            if (extents.WouldSplitAField(paragraphs[i]))
                throw new ArgumentException(
                    "The range crosses a field that begins or ends outside it, so deleting it "
                    + "would leave the field without one of its markers. Delete the field with "
                    + "word_field, or choose a range that does not cut through one.");
    }

    /// <summary>
    ///     Truncates text in the end paragraph up to the specified position.
    /// </summary>
    /// <param name="para">The paragraph.</param>
    /// <param name="endCharIndex">The character index up to which to truncate.</param>
    private static void TruncateEndParagraph(WordParagraph para, int endCharIndex)
    {
        // Everything from the start of the paragraph up to endCharIndex is inside the range. The
        // previous version only shortened the first run, and did nothing at all once endCharIndex
        // reached past that run, leaving the deleted text in place.
        DeleteSpan(para, 0, endCharIndex);
    }

    private sealed record DeleteRangeParameters(
        int StartParagraphIndex,
        int StartCharIndex,
        int EndParagraphIndex,
        int EndCharIndex);
}
