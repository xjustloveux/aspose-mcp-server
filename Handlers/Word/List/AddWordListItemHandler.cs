using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using static Aspose.Words.ConvertUtil;
using WordParagraph = Aspose.Words.Paragraph;
using WordRun = Aspose.Words.Run;

namespace AsposeMcpServer.Handlers.Word.List;

/// <summary>
///     Handler for adding a single list item to Word documents.
/// </summary>
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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var text = parameters.GetRequired<string>("text");
        var styleName = parameters.GetRequired<string>("styleName");
        var listLevel = parameters.GetOptional("listLevel", 0);
        var applyStyleIndent = parameters.GetOptional("applyStyleIndent", true);

        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("text parameter is required for add_item operation");
        if (string.IsNullOrEmpty(styleName))
            throw new ArgumentException("styleName parameter is required for add_item operation");

        var doc = context.Document;
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        var style = doc.Styles[styleName];
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
            throw new ArgumentException($"Style '{styleName}' not found. {suggestions}");
        }

        var para = new WordParagraph(doc)
        {
            ParagraphFormat = { StyleName = styleName }
        };

        if (!applyStyleIndent && listLevel > 0) para.ParagraphFormat.LeftIndent = InchToPoint(0.5 * listLevel);

        var run = new WordRun(doc, text);
        para.AppendChild(run);
        builder.CurrentParagraph.ParentNode.AppendChild(para);

        MarkModified(context);

        var result = "List item added successfully\n";
        result += $"Style: {styleName}\n";
        result += $"Level: {listLevel}\n";

        if (applyStyleIndent)
            result += "Indent: Using style-defined indent (recommended)";
        else if (listLevel > 0)
            result += $"Indent: Manually set ({listLevel * 36} points)";

        return Success(result);
    }
}
