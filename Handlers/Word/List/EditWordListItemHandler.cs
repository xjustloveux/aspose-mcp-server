using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using static Aspose.Words.ConvertUtil;
using WordParagraph = Aspose.Words.Paragraph;
using WordRun = Aspose.Words.Run;

namespace AsposeMcpServer.Handlers.Word.List;

/// <summary>
///     Handler for editing list items in Word documents.
/// </summary>
public class EditWordListItemHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "edit_item";

    /// <summary>
    ///     Edits the text and level of a list item.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: paragraphIndex, text.
    ///     Optional: level
    /// </param>
    /// <returns>Success message with edit details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var paragraphIndex = parameters.GetRequired<int>("paragraphIndex");
        var text = parameters.GetRequired<string>("text");
        var level = parameters.GetOptional<int?>("level");

        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("text parameter is required for edit_item operation");

        var doc = context.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        if (paragraphIndex < 0 || paragraphIndex >= paragraphs.Count)
            throw new ArgumentException(
                $"Paragraph index {paragraphIndex} is out of range (document has {paragraphs.Count} paragraphs)");

        if (paragraphs[paragraphIndex] is not WordParagraph para)
            throw new InvalidOperationException($"Unable to get paragraph at index {paragraphIndex}");

        para.Runs.Clear();
        var run = new WordRun(doc, text);
        para.AppendChild(run);

        if (level is >= 0 and <= 8) para.ParagraphFormat.LeftIndent = InchToPoint(0.5 * level.Value);

        MarkModified(context);

        var result = "List item edited successfully\n";
        result += $"Paragraph index: {paragraphIndex}\n";
        result += $"New text: {text}";
        if (level.HasValue) result += $"\nLevel: {level.Value}";

        return Success(result);
    }
}
