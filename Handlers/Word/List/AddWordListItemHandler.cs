using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;
using static Aspose.Words.ConvertUtil;
using WordParagraph = Aspose.Words.Paragraph;
using WordRun = Aspose.Words.Run;

namespace AsposeMcpServer.Handlers.Word.List;

/// <summary>
///     Handler for adding a single list item to Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddWordListItemHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add_item";

    /// <summary>
    ///     Adds a single list item with the specified style.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: text, styleName.
    ///     Optional: listLevel, applyStyleIndent
    /// </param>
    /// <returns>Success message with item creation details.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractAddListItemParameters(parameters);

        if (string.IsNullOrEmpty(p.Text))
            throw new ArgumentException("text parameter is required for add_item operation");
        if (string.IsNullOrEmpty(p.StyleName))
            throw new ArgumentException("styleName parameter is required for add_item operation");

        var doc = context.Document;
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        var style = doc.Styles[p.StyleName];
        if (style == null)
        {
            var commonStyles = new[]
            {
                "List Paragraph", "List Bullet", "List Number", "Heading 1", "Heading 2", "Heading 3", "Heading 4"
            };
            var availableCommon = commonStyles.Where(s => doc.Styles[s] != null).Take(3).ToList();
            var suggestions = availableCommon.Count > 0
                ? $"Common available styles: {string.Join(", ", availableCommon.Select(s => $"'{s}'"))}"
                : "Use word_get_styles tool to view available styles";
            throw new ArgumentException($"Style '{p.StyleName}' not found. {suggestions}");
        }

        var para = new WordParagraph(doc)
        {
            ParagraphFormat = { StyleName = p.StyleName }
        };

        if (p is { ApplyStyleIndent: false, ListLevel: > 0 })
            para.ParagraphFormat.LeftIndent = InchToPoint(0.5 * p.ListLevel);

        var run = new WordRun(doc, p.Text);
        para.AppendChild(run);
        builder.CurrentParagraph.ParentNode.AppendChild(para);

        MarkModified(context);

        var result = "List item added successfully\n";
        result += $"Style: {p.StyleName}\n";
        result += $"Level: {p.ListLevel}\n";

        if (p.ApplyStyleIndent)
            result += "Indent: Using style-defined indent (recommended)";
        else if (p.ListLevel > 0)
            result += $"Indent: Manually set ({p.ListLevel * 36} points)";

        return new SuccessResult { Message = result };
    }

    private static AddListItemParameters ExtractAddListItemParameters(OperationParameters parameters)
    {
        return new AddListItemParameters(
            parameters.GetRequired<string>("text"),
            parameters.GetRequired<string>("styleName"),
            parameters.GetOptional("listLevel", 0),
            parameters.GetOptional("applyStyleIndent", true));
    }

    private sealed record AddListItemParameters(
        string Text,
        string StyleName,
        int ListLevel,
        bool ApplyStyleIndent);
}
