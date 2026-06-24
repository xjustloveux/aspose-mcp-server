using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;
using static Aspose.Words.ConvertUtil;
using WordRun = Aspose.Words.Run;

namespace AsposeMcpServer.Handlers.Word.List;

/// <summary>
///     Handler for editing list items in Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractEditListItemParameters(parameters);

        if (string.IsNullOrEmpty(p.Text))
            throw new ArgumentException("text parameter is required for edit_item operation");

        var doc = context.Document;

        var para = ParagraphResolver.Resolve(doc, ParagraphAddress.From(parameters, p.ParagraphIndex)).Paragraph;

        para.Runs.Clear();
        var run = new WordRun(doc, p.Text);
        para.AppendChild(run);

        if (p.Level is >= 0 and <= 8) para.ParagraphFormat.LeftIndent = InchToPoint(0.5 * p.Level.Value);

        MarkModified(context);

        var result = "List item edited successfully\n";
        result += $"Paragraph index: {p.ParagraphIndex}\n";
        result += $"New text: {p.Text}";
        if (p.Level.HasValue) result += $"\nLevel: {p.Level.Value}";

        return new SuccessResult { Message = result };
    }

    private static EditListItemParameters ExtractEditListItemParameters(OperationParameters parameters)
    {
        return new EditListItemParameters(
            parameters.GetRequired<int>("paragraphIndex"),
            parameters.GetRequired<string>("text"),
            parameters.GetOptional<int?>("level"));
    }

    private sealed record EditListItemParameters(
        int ParagraphIndex,
        string Text,
        int? Level);
}
