using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Text;

/// <summary>
///     Handler for deleting text range in Word documents.
/// </summary>
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
    ///     Optional: sectionIndex.
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when indices are out of range.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var startParagraphIndex = parameters.GetRequired<int>("startParagraphIndex");
        var startCharIndex = parameters.GetRequired<int>("startCharIndex");
        var endParagraphIndex = parameters.GetRequired<int>("endParagraphIndex");
        var endCharIndex = parameters.GetRequired<int>("endCharIndex");
        var sectionIndex = parameters.GetOptional("sectionIndex", 0);

        var doc = context.Document;

        ValidateSectionIndex(doc, sectionIndex);
        var section = doc.Sections[sectionIndex];
        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();

        ValidateParagraphIndices(paragraphs, startParagraphIndex, endParagraphIndex);

        var startPara = paragraphs[startParagraphIndex];
        var endPara = paragraphs[endParagraphIndex];

        if (startParagraphIndex == endParagraphIndex)
            DeleteWithinSameParagraph(startPara, startCharIndex, endCharIndex);
        else
            DeleteAcrossParagraphs(paragraphs, startPara, endPara, startParagraphIndex, endParagraphIndex,
                startCharIndex, endCharIndex);

        MarkModified(context);

        return Success("Text range deleted.");
    }

    /// <summary>
    ///     Validates the section index is within range.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="sectionIndex">The section index to validate.</param>
    /// <exception cref="ArgumentException">Thrown when section index is out of range.</exception>
    private static void ValidateSectionIndex(Document doc, int sectionIndex)
    {
        if (sectionIndex < 0 || sectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
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
    ///     Deletes text within a single paragraph.
    /// </summary>
    /// <param name="para">The paragraph.</param>
    /// <param name="startCharIndex">The start character index.</param>
    /// <param name="endCharIndex">The end character index.</param>
    private static void DeleteWithinSameParagraph(WordParagraph para, int startCharIndex, int endCharIndex)
    {
        var runs = para.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var (startRunIndex, startRunCharIndex, endRunIndex, endRunCharIndex) =
            FindRunRange(runs, startCharIndex, endCharIndex);

        if (startRunIndex < 0 || endRunIndex < 0) return;

        if (startRunIndex == endRunIndex)
            DeleteWithinSameRun(runs[startRunIndex], startRunCharIndex, endRunCharIndex);
        else
            DeleteAcrossRuns(runs, startRunIndex, startRunCharIndex, endRunIndex, endRunCharIndex);
    }

    /// <summary>
    ///     Finds the run indices and character positions for the deletion range.
    /// </summary>
    /// <param name="runs">The list of runs.</param>
    /// <param name="startCharIndex">The start character index.</param>
    /// <param name="endCharIndex">The end character index.</param>
    /// <returns>Tuple of start run index, start char index, end run index, end char index.</returns>
    private static (int startRunIndex, int startRunCharIndex, int endRunIndex, int endRunCharIndex)
        FindRunRange(List<Run> runs, int startCharIndex, int endCharIndex)
    {
        var totalChars = 0;
        int startRunIndex = -1, endRunIndex = -1;
        int startRunCharIndex = 0, endRunCharIndex = 0;

        for (var i = 0; i < runs.Count; i++)
        {
            var runLength = runs[i].Text.Length;

            if (startRunIndex == -1 && totalChars + runLength > startCharIndex)
            {
                startRunIndex = i;
                startRunCharIndex = startCharIndex - totalChars;
            }

            if (totalChars + runLength > endCharIndex)
            {
                endRunIndex = i;
                endRunCharIndex = endCharIndex - totalChars;
                break;
            }

            totalChars += runLength;
        }

        return (startRunIndex, startRunCharIndex, endRunIndex, endRunCharIndex);
    }

    /// <summary>
    ///     Deletes text within a single run.
    /// </summary>
    /// <param name="run">The run.</param>
    /// <param name="startCharIndex">The start character index within the run.</param>
    /// <param name="endCharIndex">The end character index within the run.</param>
    private static void DeleteWithinSameRun(Run run, int startCharIndex, int endCharIndex)
    {
        run.Text = run.Text.Remove(startCharIndex, endCharIndex - startCharIndex);
    }

    /// <summary>
    ///     Deletes text across multiple runs.
    /// </summary>
    /// <param name="runs">The list of runs.</param>
    /// <param name="startRunIndex">The start run index.</param>
    /// <param name="startRunCharIndex">The start character index within the start run.</param>
    /// <param name="endRunIndex">The end run index.</param>
    /// <param name="endRunCharIndex">The end character index within the end run.</param>
    private static void DeleteAcrossRuns(List<Run> runs, int startRunIndex, int startRunCharIndex,
        int endRunIndex, int endRunCharIndex)
    {
        var startRun = runs[startRunIndex];
        startRun.Text = startRun.Text.Substring(0, startRunCharIndex);

        for (var i = startRunIndex + 1; i < endRunIndex; i++)
            runs[i].Remove();

        if (endRunIndex < runs.Count)
        {
            var endRun = runs[endRunIndex];
            endRun.Text = endRun.Text.Substring(endRunCharIndex);
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
        var runs = para.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var lastRun = runs.LastOrDefault();
        if (lastRun != null && lastRun.Text.Length > startCharIndex)
            lastRun.Text = lastRun.Text.Substring(0, startCharIndex);
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
    ///     Truncates text in the end paragraph up to the specified position.
    /// </summary>
    /// <param name="para">The paragraph.</param>
    /// <param name="endCharIndex">The character index up to which to truncate.</param>
    private static void TruncateEndParagraph(WordParagraph para, int endCharIndex)
    {
        var runs = para.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        if (runs.Count > 0 && endCharIndex < runs[0].Text.Length)
        {
            runs[0].Text = runs[0].Text.Substring(endCharIndex);
            for (var i = 1; i < runs.Count; i++)
                runs[i].Remove();
        }
    }
}
