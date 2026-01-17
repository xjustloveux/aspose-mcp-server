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
    public override string
        Execute(OperationContext<Document> context,
            OperationParameters parameters) // NOSONAR S3776 - Linear validation and conversion
    {
        var p = ExtractConvertToListParameters(parameters);

        var doc = context.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();

        if (p.StartIndex < 0 || p.StartIndex >= paragraphs.Count)
            throw new ArgumentException(
                $"Start paragraph index {p.StartIndex} is out of range (document has {paragraphs.Count} paragraphs)");

        if (p.EndIndex < 0 || p.EndIndex >= paragraphs.Count)
            throw new ArgumentException(
                $"End paragraph index {p.EndIndex} is out of range (document has {paragraphs.Count} paragraphs)");

        if (p.StartIndex > p.EndIndex)
            throw new ArgumentException(
                $"Start index ({p.StartIndex}) must be less than or equal to end index ({p.EndIndex})");

        var list = doc.Lists.Add(p.ListType == "number"
            ? ListTemplate.NumberDefault
            : ListTemplate.BulletDefault);

        if (p.ListType == "number")
        {
            var numStyle = p.NumberFormat.ToLower() switch
            {
                "roman" => NumberStyle.UppercaseRoman,
                "letter" => NumberStyle.UppercaseLetter,
                _ => NumberStyle.Arabic
            };

            foreach (var level in list.ListLevels) level.NumberStyle = numStyle;
        }

        var convertedCount = 0;
        var skippedCount = 0;
        for (var i = p.StartIndex; i <= p.EndIndex; i++)
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
        result += $"Range: paragraph {p.StartIndex} to {p.EndIndex}\n";
        result += $"List type: {p.ListType}\n";
        if (p.ListType == "number") result += $"Number format: {p.NumberFormat}\n";
        result += $"Converted: {convertedCount} paragraphs";
        if (skippedCount > 0) result += $"\nSkipped: {skippedCount} paragraphs (already list items or empty)";

        return Success(result);
    }

    private static ConvertToListParameters ExtractConvertToListParameters(OperationParameters parameters)
    {
        return new ConvertToListParameters(
            parameters.GetRequired<int>("startParagraphIndex"),
            parameters.GetRequired<int>("endParagraphIndex"),
            parameters.GetOptional("listType", "bullet"),
            parameters.GetOptional("numberFormat", "arabic"));
    }

    private sealed record ConvertToListParameters(
        int StartIndex,
        int EndIndex,
        string ListType,
        string NumberFormat);
}
