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
        var p = ExtractAddListParameters(parameters);

        if (p.Items.Count == 0)
            throw new ArgumentException("items parameter is required and cannot be empty");

        var parsedItems = WordListHelper.ParseItems(p.Items);
        var doc = context.Document;
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        var (list, isContinuing) = GetOrCreateList(doc, p.ContinuePrevious, p.ListType, p.BulletChar, p.NumberFormat);

        WriteListItems(builder, list, parsedItems);

        builder.ListFormat.RemoveNumbers();
        MarkModified(context);

        return Success(BuildResultMessage(isContinuing, p.ListType, p.BulletChar, p.NumberFormat, list,
            parsedItems.Count));
    }

    /// <summary>
    ///     Extracts add list parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted add list parameters.</returns>
    private static AddListParameters ExtractAddListParameters(OperationParameters parameters)
    {
        return new AddListParameters(
            parameters.GetRequired<JsonArray>("items"),
            parameters.GetOptional("listType", "bullet"),
            parameters.GetOptional("bulletChar", "\u2022"),
            parameters.GetOptional("numberFormat", "arabic"),
            parameters.GetOptional("continuePrevious", false)
        );
    }

    /// <summary>
    ///     Gets an existing list to continue or creates a new one.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="continuePrevious">Whether to continue a previous list.</param>
    /// <param name="listType">The list type (bullet, number, custom).</param>
    /// <param name="bulletChar">The bullet character for custom lists.</param>
    /// <param name="numberFormat">The number format for numbered lists.</param>
    /// <returns>A tuple containing the list and whether it's continuing a previous list.</returns>
    private static (WordList list, bool isContinuing) GetOrCreateList(Document doc, bool continuePrevious,
        string listType, string bulletChar, string numberFormat)
    {
        if (continuePrevious && doc.Lists.Count > 0)
        {
            var existingList = FindExistingList(doc);
            if (existingList != null)
                return (existingList, true);
        }

        var newList = CreateNewList(doc, listType, bulletChar, numberFormat);
        return (newList, false);
    }

    /// <summary>
    ///     Finds the most recent list in the document.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <returns>The most recent list or null if none found.</returns>
    private static WordList? FindExistingList(Document doc)
    {
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
        for (var i = paragraphs.Count - 1; i >= 0; i--)
            if (paragraphs[i].ListFormat is { IsListItem: true, List: not null })
                return paragraphs[i].ListFormat.List;

        return null;
    }

    /// <summary>
    ///     Creates a new list with the specified type and formatting.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="listType">The list type.</param>
    /// <param name="bulletChar">The bullet character.</param>
    /// <param name="numberFormat">The number format.</param>
    /// <returns>The newly created list.</returns>
    private static WordList CreateNewList(Document doc, string listType, string bulletChar, string numberFormat)
    {
        var list = doc.Lists.Add(listType == "number" ? ListTemplate.NumberDefault : ListTemplate.BulletDefault);

        if (listType == "custom" && !string.IsNullOrEmpty(bulletChar))
        {
            list.ListLevels[0].NumberFormat = bulletChar;
            list.ListLevels[0].NumberStyle = NumberStyle.Bullet;
        }
        else if (listType == "number")
        {
            var numStyle = ParseNumberStyle(numberFormat);
            foreach (var level in list.ListLevels) level.NumberStyle = numStyle;
        }

        return list;
    }

    /// <summary>
    ///     Parses a number format string to NumberStyle enum.
    /// </summary>
    /// <param name="numberFormat">The number format string.</param>
    /// <returns>The corresponding NumberStyle enum value.</returns>
    private static NumberStyle ParseNumberStyle(string numberFormat)
    {
        return numberFormat.ToLower() switch
        {
            "roman" => NumberStyle.UppercaseRoman,
            "letter" => NumberStyle.UppercaseLetter,
            _ => NumberStyle.Arabic
        };
    }

    /// <summary>
    ///     Writes list items to the document.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="list">The list to write items to.</param>
    /// <param name="parsedItems">The parsed list items.</param>
    private static void WriteListItems(DocumentBuilder builder, WordList list,
        List<(string text, int level)> parsedItems)
    {
        foreach (var item in parsedItems)
        {
            builder.ListFormat.List = list;
            builder.ListFormat.ListLevelNumber = Math.Min(item.level, 8);
            builder.Writeln(item.text);
        }
    }

    /// <summary>
    ///     Builds the result message for a successful list addition.
    /// </summary>
    /// <param name="isContinuing">Whether the list continues a previous one.</param>
    /// <param name="listType">The list type.</param>
    /// <param name="bulletChar">The bullet character.</param>
    /// <param name="numberFormat">The number format.</param>
    /// <param name="list">The created list.</param>
    /// <param name="itemCount">The number of items added.</param>
    /// <returns>The result message.</returns>
    private static string BuildResultMessage(bool isContinuing, string listType, string bulletChar,
        string numberFormat, WordList list, int itemCount)
    {
        var result = isContinuing ? "List items added (continuing previous list)\n" : "List added successfully\n";

        if (isContinuing)
        {
            result += $"Continued from list ID: {list.ListId}\n";
        }
        else
        {
            result += $"Type: {listType}\n";
            if (listType == "custom") result += $"Bullet character: {bulletChar}\n";
            if (listType == "number") result += $"Number format: {numberFormat}\n";
        }

        result += $"Item count: {itemCount}";
        return result;
    }

    /// <summary>
    ///     Record to hold add list parameters.
    /// </summary>
    private record AddListParameters(
        JsonArray Items,
        string ListType,
        string BulletChar,
        string NumberFormat,
        bool ContinuePrevious);
}
