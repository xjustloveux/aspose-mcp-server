using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Text;

/// <summary>
///     Represents a run index and character index within that run.
/// </summary>
/// <param name="RunIndex">The index of the run.</param>
/// <param name="CharacterIndex">The character index within the run.</param>
internal record RunPosition(int RunIndex, int CharacterIndex);

/// <summary>
///     Handler for inserting text at a specific position in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class InsertAtPositionWordTextHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "insert";

    /// <summary>
    ///     Inserts text at a specific paragraph and character position.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: insertParagraphIndex, charIndex, text.
    ///     Optional: insertBefore.
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or indices are out of range.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractInsertAtPositionParameters(parameters);

        var doc = context.Document;

        var para = ParagraphResolver.Resolve(doc, ParagraphAddress.From(parameters, p.ParagraphIndex)).Paragraph;

        InsertText(doc, para, p.CharIndex, p.Text, p.InsertBefore);

        MarkModified(context);

        return new SuccessResult { Message = "Text inserted at position." };
    }

    private static InsertAtPositionParameters ExtractInsertAtPositionParameters(OperationParameters parameters)
    {
        return new InsertAtPositionParameters(
            parameters.GetRequired<int>("insertParagraphIndex"),
            parameters.GetRequired<int>("charIndex"),
            parameters.GetRequired<string>("text"),
            parameters.GetOptional("insertBefore", false));
    }

    /// <summary>
    ///     Inserts text at the specified position.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="para">The target paragraph.</param>
    /// <param name="charIndex">The character index.</param>
    /// <param name="text">The text to insert.</param>
    /// <param name="insertBefore">Whether to insert before the position.</param>
    private static void InsertText(Document doc, WordParagraph? para, int charIndex,
        string text, bool insertBefore)
    {
        if (para == null)
            throw new ArgumentNullException(nameof(para), "Paragraph cannot be null");

        var runs = para.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        var runPosition = FindTargetRunPosition(runs, charIndex);

        if (runPosition.RunIndex == -1)
        {
            InsertUsingBuilder(doc, para, text);
            return;
        }

        var targetRun = runs[runPosition.RunIndex];
        var enclosingField = FieldBoundaryHelper.GetEnclosingField(targetRun);
        if (enclosingField != null)
        {
            InsertAtFieldBoundary(doc, enclosingField, text, insertBefore);
            return;
        }

        InsertIntoRun(targetRun, runPosition.CharacterIndex, text);
    }

    /// <summary>
    ///     Inserts text at a field boundary instead of inside the field's node range, preventing
    ///     field-code/result corruption. When insertBefore is true the text is placed before the
    ///     field's start; otherwise it is placed after the field's end.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="field">The field whose boundary anchors the insertion.</param>
    /// <param name="text">The text to insert.</param>
    /// <param name="insertBefore">Whether to insert before the field rather than after it.</param>
    private static void InsertAtFieldBoundary(Document doc, Aspose.Words.Fields.Field field, string text,
        bool insertBefore)
    {
        var builder = new DocumentBuilder(doc);
        if (insertBefore)
            builder.MoveTo(field.Start);
        else
            FieldBoundaryHelper.MoveToAfterField(builder, field);

        builder.Write(text);
    }

    /// <summary>
    ///     Finds the run and character position for insertion.
    /// </summary>
    /// <param name="runs">The list of runs.</param>
    /// <param name="charIndex">The target character index.</param>
    /// <returns>A RunPosition containing run index and character index within the run.</returns>
    private static RunPosition FindTargetRunPosition(List<Run> runs, int charIndex)
    {
        var totalChars = 0;

        for (var i = 0; i < runs.Count; i++)
        {
            var runLength = runs[i].Text.Length;
            if (totalChars + runLength >= charIndex)
                return new RunPosition(i, charIndex - totalChars);
            totalChars += runLength;
        }

        return new RunPosition(-1, 0);
    }

    /// <summary>
    ///     Inserts text at the end of the paragraph using DocumentBuilder when no run matches the
    ///     requested character position (e.g. an empty paragraph).
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="para">The target paragraph.</param>
    /// <param name="text">The text to insert.</param>
    private static void InsertUsingBuilder(Document doc, WordParagraph? para, string text)
    {
        if (para == null)
            throw new ArgumentNullException(nameof(para), "Paragraph cannot be null");

        var builder = new DocumentBuilder(doc);
        builder.MoveTo(para);
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

    private sealed record InsertAtPositionParameters(
        int ParagraphIndex,
        int CharIndex,
        string Text,
        bool InsertBefore);
}
