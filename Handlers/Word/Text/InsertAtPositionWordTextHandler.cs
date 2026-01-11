using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Text;

/// <summary>
///     Handler for inserting text at a specific position in Word documents.
/// </summary>
public class InsertAtPositionWordTextHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "insert_at_position";

    /// <summary>
    ///     Inserts text at a specific paragraph and character position.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: insertParagraphIndex, charIndex, text.
    ///     Optional: sectionIndex, insertBefore.
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or indices are out of range.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var paragraphIndex = parameters.GetRequired<int>("insertParagraphIndex");
        var charIndex = parameters.GetRequired<int>("charIndex");
        var text = parameters.GetRequired<string>("text");
        var sectionIndex = parameters.GetOptional("sectionIndex", 0);
        var insertBefore = parameters.GetOptional("insertBefore", false);

        var doc = context.Document;

        ValidateSectionIndex(doc, sectionIndex);
        var section = doc.Sections[sectionIndex];
        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();

        ValidateParagraphIndex(paragraphs, paragraphIndex);
        var para = paragraphs[paragraphIndex];

        InsertText(doc, para, paragraphIndex, charIndex, text, insertBefore);

        MarkModified(context);

        return Success("Text inserted at position.");
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
    ///     Validates the paragraph index is within range.
    /// </summary>
    /// <param name="paragraphs">The list of paragraphs.</param>
    /// <param name="paragraphIndex">The paragraph index to validate.</param>
    /// <exception cref="ArgumentException">Thrown when paragraph index is out of range.</exception>
    private static void ValidateParagraphIndex(List<WordParagraph> paragraphs, int paragraphIndex)
    {
        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
            throw new ArgumentException($"paragraphIndex must be between 0 and {paragraphs.Count - 1}");
    }

    /// <summary>
    ///     Inserts text at the specified position.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="para">The target paragraph.</param>
    /// <param name="paragraphIndex">The paragraph index.</param>
    /// <param name="charIndex">The character index.</param>
    /// <param name="text">The text to insert.</param>
    /// <param name="insertBefore">Whether to insert before the position.</param>
    private static void InsertText(Document doc, WordParagraph? para, int paragraphIndex, int charIndex,
        string text, bool insertBefore)
    {
        if (para == null)
            throw new ArgumentNullException(nameof(para), "Paragraph cannot be null");

        var runs = para.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var (targetRunIndex, targetRunCharIndex) = FindTargetRunPosition(runs, charIndex);

        if (targetRunIndex == -1)
            InsertUsingBuilder(doc, para, paragraphIndex, text, insertBefore);
        else
            InsertIntoRun(runs[targetRunIndex], targetRunCharIndex, text);
    }

    /// <summary>
    ///     Finds the run and character position for insertion.
    /// </summary>
    /// <param name="runs">The list of runs.</param>
    /// <param name="charIndex">The target character index.</param>
    /// <returns>Tuple of run index and character index within the run.</returns>
    private static (int runIndex, int charIndex) FindTargetRunPosition(List<Run> runs, int charIndex)
    {
        var totalChars = 0;

        for (var i = 0; i < runs.Count; i++)
        {
            var runLength = runs[i].Text.Length;
            if (totalChars + runLength >= charIndex)
                return (i, charIndex - totalChars);
            totalChars += runLength;
        }

        return (-1, 0);
    }

    /// <summary>
    ///     Inserts text using DocumentBuilder when no suitable run is found.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="para">The target paragraph.</param>
    /// <param name="paragraphIndex">The paragraph index.</param>
    /// <param name="text">The text to insert.</param>
    /// <param name="insertBefore">Whether to insert before the position.</param>
    private static void InsertUsingBuilder(Document doc, WordParagraph? para, int paragraphIndex,
        string text, bool insertBefore)
    {
        if (para == null)
            throw new ArgumentNullException(nameof(para), "Paragraph cannot be null");

        var builder = new DocumentBuilder(doc);
        builder.MoveTo(para);

        if (!insertBefore)
            builder.MoveToParagraph(paragraphIndex, para.GetText().Length);

        builder.Write(text);
    }

    /// <summary>
    ///     Inserts text directly into an existing run.
    /// </summary>
    /// <param name="run">The target run.</param>
    /// <param name="charIndex">The character index within the run.</param>
    /// <param name="text">The text to insert.</param>
    private static void InsertIntoRun(Run run, int charIndex, string text)
    {
        run.Text = run.Text.Insert(charIndex, text);
    }
}
