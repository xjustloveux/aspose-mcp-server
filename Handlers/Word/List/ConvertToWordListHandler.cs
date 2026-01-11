using Aspose.Words;
using Aspose.Words.Lists;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.List;

/// <summary>
///     Handler for converting paragraphs to lists in Word documents.
/// </summary>
public class ConvertToWordListHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "convert_to_list";

    /// <summary>
    ///     Converts a range of paragraphs into a list.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: startParagraphIndex, endParagraphIndex.
    ///     Optional: listType, numberFormat
    /// </param>
    /// <returns>Success message with conversion details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var startIndex = parameters.GetRequired<int>("startParagraphIndex");
        var endIndex = parameters.GetRequired<int>("endParagraphIndex");
        var listType = parameters.GetOptional("listType", "bullet");
        var numberFormat = parameters.GetOptional("numberFormat", "arabic");

        var doc = context.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();

        if (startIndex < 0 || startIndex >= paragraphs.Count)
            throw new ArgumentException(
                $"Start paragraph index {startIndex} is out of range (document has {paragraphs.Count} paragraphs)");

        if (endIndex < 0 || endIndex >= paragraphs.Count)
            throw new ArgumentException(
                $"End paragraph index {endIndex} is out of range (document has {paragraphs.Count} paragraphs)");

        if (startIndex > endIndex)
            throw new ArgumentException(
                $"Start index ({startIndex}) must be less than or equal to end index ({endIndex})");

        var list = doc.Lists.Add(listType == "number"
            ? ListTemplate.NumberDefault
            : ListTemplate.BulletDefault);

        if (listType == "number")
        {
            var numStyle = numberFormat.ToLower() switch
            {
                "roman" => NumberStyle.UppercaseRoman,
                "letter" => NumberStyle.UppercaseLetter,
                _ => NumberStyle.Arabic
            };

            foreach (var level in list.ListLevels) level.NumberStyle = numStyle;
        }

        var convertedCount = 0;
        var skippedCount = 0;
        for (var i = startIndex; i <= endIndex; i++)
        {
            var para = paragraphs[i];

            if (para.ListFormat.IsListItem)
            {
                skippedCount++;
                continue;
            }

            var text = para.ToString(SaveFormat.Text).Trim();
            if (string.IsNullOrEmpty(text))
            {
                skippedCount++;
                continue;
            }

            para.ListFormat.List = list;
            para.ListFormat.ListLevelNumber = 0;
            convertedCount++;
        }

        MarkModified(context);

        var result = "Paragraphs converted to list successfully\n";
        result += $"Range: paragraph {startIndex} to {endIndex}\n";
        result += $"List type: {listType}\n";
        if (listType == "number") result += $"Number format: {numberFormat}\n";
        result += $"Converted: {convertedCount} paragraphs";
        if (skippedCount > 0) result += $"\nSkipped: {skippedCount} paragraphs (already list items or empty)";

        return Success(result);
    }
}
