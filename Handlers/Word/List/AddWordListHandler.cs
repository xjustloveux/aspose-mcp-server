using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Lists;
using AsposeMcpServer.Core.Handlers;
using WordList = Aspose.Words.Lists.List;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.List;

/// <summary>
///     Handler for adding lists to Word documents.
/// </summary>
public class AddWordListHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add_list";

    /// <summary>
    ///     Adds a new list with the specified items and formatting.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: items (JsonArray).
    ///     Optional: listType, bulletChar, numberFormat, continuePrevious
    /// </param>
    /// <returns>Success message with list creation details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var items = parameters.GetRequired<JsonArray>("items");
        var listType = parameters.GetOptional("listType", "bullet");
        var bulletChar = parameters.GetOptional("bulletChar", "•");
        var numberFormat = parameters.GetOptional("numberFormat", "arabic");
        var continuePrevious = parameters.GetOptional("continuePrevious", false);

        if (items.Count == 0)
            throw new ArgumentException("items parameter is required and cannot be empty");

        var parsedItems = WordListHelper.ParseItems(items);
        var doc = context.Document;
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        WordList? list = null;
        var isContinuing = false;

        if (continuePrevious && doc.Lists.Count > 0)
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
            for (var i = paragraphs.Count - 1; i >= 0; i--)
                if (paragraphs[i].ListFormat is { IsListItem: true, List: not null })
                {
                    list = paragraphs[i].ListFormat.List;
                    isContinuing = true;
                    break;
                }
        }

        if (list == null)
        {
            list = doc.Lists.Add(listType == "number"
                ? ListTemplate.NumberDefault
                : ListTemplate.BulletDefault);

            if (listType == "custom" && !string.IsNullOrEmpty(bulletChar))
            {
                list.ListLevels[0].NumberFormat = bulletChar;
                list.ListLevels[0].NumberStyle = NumberStyle.Bullet;
            }
            else if (listType == "number")
            {
                var numStyle = numberFormat.ToLower() switch
                {
                    "roman" => NumberStyle.UppercaseRoman,
                    "letter" => NumberStyle.UppercaseLetter,
                    _ => NumberStyle.Arabic
                };

                foreach (var level in list.ListLevels) level.NumberStyle = numStyle;
            }
        }

        foreach (var item in parsedItems)
        {
            builder.ListFormat.List = list;
            builder.ListFormat.ListLevelNumber = Math.Min(item.level, 8);
            builder.Writeln(item.text);
        }

        builder.ListFormat.RemoveNumbers();
        MarkModified(context);

        var result = isContinuing
            ? "List items added (continuing previous list)\n"
            : "List added successfully\n";
        if (!isContinuing)
        {
            result += $"Type: {listType}\n";
            if (listType == "custom") result += $"Bullet character: {bulletChar}\n";
            if (listType == "number") result += $"Number format: {numberFormat}\n";
        }
        else
        {
            result += $"Continued from list ID: {list.ListId}\n";
        }

        result += $"Item count: {parsedItems.Count}";

        return Success(result);
    }
}
