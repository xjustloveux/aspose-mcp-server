using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Lists;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;
using static Aspose.Words.ConvertUtil;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for list operations in Word documents
///     Merges: WordAddListTool, WordAddListItemTool, WordDeleteListItemTool, WordEditListItemTool,
///     WordSetListFormatTool, WordGetListFormatTool
/// </summary>
[McpServerToolType]
public class WordListTool
{
    /// <summary>
    ///     Session manager for document session operations
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the WordListTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    public WordListTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "word_list")]
    [Description(
        @"Manage lists in Word documents. Supports 8 operations: add_list, add_item, delete_item, edit_item, set_format, get_format, restart_numbering, convert_to_list.

Usage examples:
- Add bullet list: word_list(operation='add_list', path='doc.docx', items=['Item 1', 'Item 2', 'Item 3'])
- Add numbered list: word_list(operation='add_list', path='doc.docx', items=['First', 'Second'], listType='number')
- Add list item: word_list(operation='add_item', path='doc.docx', text='New item', styleName='Heading 4')
- Delete list item: word_list(operation='delete_item', path='doc.docx', paragraphIndex=0)
- Edit list item: word_list(operation='edit_item', path='doc.docx', paragraphIndex=0, text='Updated text')
- Get list format: word_list(operation='get_format', path='doc.docx', paragraphIndex=0)
- Restart numbering: word_list(operation='restart_numbering', path='doc.docx', paragraphIndex=2, startAt=1)
- Convert to list: word_list(operation='convert_to_list', path='doc.docx', startParagraphIndex=0, endParagraphIndex=5)")]
    public string Execute(
        [Description(
            "Operation: add_list, add_item, delete_item, edit_item, set_format, get_format, restart_numbering, convert_to_list")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("List items for add_list operation (string array or object array with text/level)")]
        JsonArray? items = null,
        [Description("List type: bullet, number, custom (default: bullet)")]
        string listType = "bullet",
        [Description("Custom bullet character (for custom type)")]
        string bulletChar = "•",
        [Description("Number format: arabic, roman, letter (default: arabic)")]
        string numberFormat = "arabic",
        [Description("Continue numbering from last list (default: false)")]
        bool continuePrevious = false,
        [Description("List item text content")]
        string? text = null,
        [Description("Style name for the list item")]
        string? styleName = null,
        [Description("List level (0-8)")] int listLevel = 0,
        [Description("Use style-defined indent (default: true)")]
        bool applyStyleIndent = true,
        [Description("Paragraph index (0-based)")]
        int? paragraphIndex = null,
        [Description("List level for edit (0-8)")]
        int? level = null,
        [Description("Number style: arabic, roman, letter, bullet, none")]
        string? numberStyle = null,
        [Description("Indentation level (0-8, each level = 36 points)")]
        int? indentLevel = null,
        [Description("Left indent in points")] double? leftIndent = null,
        [Description("First line indent in points")]
        double? firstLineIndent = null,
        [Description("Number to restart at (default: 1)")]
        int startAt = 1,
        [Description("Starting paragraph index (for convert_to_list)")]
        int? startParagraphIndex = null,
        [Description("Ending paragraph index (for convert_to_list)")]
        int? endParagraphIndex = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);

        return operation.ToLower() switch
        {
            "add_list" => AddList(ctx, outputPath, items, listType, bulletChar, numberFormat, continuePrevious),
            "add_item" => AddListItem(ctx, outputPath, text, styleName, listLevel, applyStyleIndent),
            "delete_item" => DeleteListItem(ctx, outputPath, paragraphIndex),
            "edit_item" => EditListItem(ctx, outputPath, paragraphIndex, text, level),
            "set_format" => SetListFormat(ctx, outputPath, paragraphIndex, numberStyle, indentLevel, leftIndent,
                firstLineIndent),
            "get_format" => GetListFormat(ctx, paragraphIndex),
            "restart_numbering" => RestartNumbering(ctx, outputPath, paragraphIndex, startAt),
            "convert_to_list" => ConvertToList(ctx, outputPath, startParagraphIndex, endParagraphIndex, listType,
                numberFormat),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a new list with the specified items and formatting.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="items">The list items as JSON array.</param>
    /// <param name="listType">The list type (bullet, number, custom).</param>
    /// <param name="bulletChar">The custom bullet character.</param>
    /// <param name="numberFormat">The number format (arabic, roman, letter).</param>
    /// <param name="continuePrevious">Whether to continue numbering from the previous list.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when items parameter is null or empty.</exception>
    private static string AddList(DocumentContext<Document> ctx, string? outputPath, JsonArray? items, string listType,
        string bulletChar, string numberFormat, bool continuePrevious)
    {
        if (items == null || items.Count == 0)
            throw new ArgumentException("items parameter is required and cannot be empty");

        var parsedItems = ParseItems(items);
        var doc = ctx.Document;
        var builder = new DocumentBuilder(doc);
        builder.MoveToDocumentEnd();

        List? list = null;
        var isContinuing = false;

        // Try to continue from previous list if requested
        if (continuePrevious && doc.Lists.Count > 0)
        {
            var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
            for (var i = paragraphs.Count - 1; i >= 0; i--)
                if (paragraphs[i].ListFormat is { IsListItem: true, List: not null })
                {
                    list = paragraphs[i].ListFormat.List;
                    isContinuing = true;
                    break;
                }
        }

        // Create new list if not continuing
        if (list == null)
        {
            list = doc.Lists.Add(listType == "number"
                ? ListTemplate.NumberDefault
                : ListTemplate.BulletDefault);

            // Configure list format for new lists only
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

        // Add list items
        foreach (var item in parsedItems)
        {
            builder.ListFormat.List = list;
            builder.ListFormat.ListLevelNumber = Math.Min(item.level, 8);
            builder.Writeln(item.text);
        }

        // Remove list formatting after adding items
        builder.ListFormat.RemoveNumbers();
        ctx.Save(outputPath);

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

        result += $"Item count: {parsedItems.Count}\n";
        result += ctx.GetOutputMessage(outputPath);

        return result;
    }

    /// <summary>
    ///     Adds a single list item with the specified style.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="text">The text content of the list item.</param>
    /// <param name="styleName">The style name for the list item.</param>
    /// <param name="listLevel">The list level (0-8).</param>
    /// <param name="applyStyleIndent">Whether to use style-defined indent.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when text or styleName is empty, or style is not found.</exception>
    private static string AddListItem(DocumentContext<Document> ctx, string? outputPath, string? text,
        string? styleName, int listLevel, bool applyStyleIndent)
    {
        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("text parameter is required for add_item operation");
        if (string.IsNullOrEmpty(styleName))
            throw new ArgumentException("styleName parameter is required for add_item operation");

        var doc = ctx.Document;
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

        var para = new Paragraph(doc)
        {
            ParagraphFormat = { StyleName = styleName }
        };

        if (!applyStyleIndent && listLevel > 0) para.ParagraphFormat.LeftIndent = InchToPoint(0.5 * listLevel);

        var run = new Run(doc, text);
        para.AppendChild(run);
        builder.CurrentParagraph.ParentNode.AppendChild(para);

        ctx.Save(outputPath);

        var result = "List item added successfully\n";
        result += $"Style: {styleName}\n";
        result += $"Level: {listLevel}\n";

        if (applyStyleIndent)
            result += "Indent: Using style-defined indent (recommended)\n";
        else if (listLevel > 0) result += $"Indent: Manually set ({listLevel * 36} points)\n";

        result += ctx.GetOutputMessage(outputPath);

        return result;
    }

    /// <summary>
    ///     Deletes a list item at the specified paragraph index.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="paragraphIndex">The zero-based paragraph index.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when paragraphIndex is not provided or out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when the paragraph cannot be accessed.</exception>
    private static string DeleteListItem(DocumentContext<Document> ctx, string? outputPath, int? paragraphIndex)
    {
        if (!paragraphIndex.HasValue)
            throw new ArgumentException("paragraphIndex parameter is required for delete_item operation");

        var doc = ctx.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
            throw new ArgumentException(
                $"Paragraph index {paragraphIndex.Value} is out of range (document has {paragraphs.Count} paragraphs)");

        if (paragraphs[paragraphIndex.Value] is not Paragraph paraToDelete)
            throw new InvalidOperationException($"Unable to get paragraph at index {paragraphIndex.Value}");

        var itemText = paraToDelete.GetText().Trim();
        var itemPreview = itemText.Length > 50 ? itemText.Substring(0, 50) + "..." : itemText;
        var isListItem = paraToDelete.ListFormat.IsListItem;
        var listInfo = isListItem ? " (list item)" : " (regular paragraph)";

        paraToDelete.Remove();
        ctx.Save(outputPath);

        var result = $"List item #{paragraphIndex.Value} deleted successfully{listInfo}\n";
        if (!string.IsNullOrEmpty(itemPreview)) result += $"Content preview: {itemPreview}\n";
        result += $"Remaining paragraphs: {doc.GetChildNodes(NodeType.Paragraph, true).Count}\n";
        result += ctx.GetOutputMessage(outputPath);

        return result;
    }

    /// <summary>
    ///     Edits the text and level of a list item.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="paragraphIndex">The zero-based paragraph index.</param>
    /// <param name="text">The new text content.</param>
    /// <param name="level">The new list level (0-8).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when paragraphIndex or text is not provided, or index is out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when the paragraph cannot be accessed.</exception>
    private static string EditListItem(DocumentContext<Document> ctx, string? outputPath, int? paragraphIndex,
        string? text, int? level)
    {
        if (!paragraphIndex.HasValue)
            throw new ArgumentException("paragraphIndex parameter is required for edit_item operation");
        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("text parameter is required for edit_item operation");

        var doc = ctx.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
            throw new ArgumentException(
                $"Paragraph index {paragraphIndex.Value} is out of range (document has {paragraphs.Count} paragraphs)");

        if (paragraphs[paragraphIndex.Value] is not Paragraph para)
            throw new InvalidOperationException($"Unable to get paragraph at index {paragraphIndex.Value}");

        para.Runs.Clear();
        var run = new Run(doc, text);
        para.AppendChild(run);

        if (level is >= 0 and <= 8) para.ParagraphFormat.LeftIndent = InchToPoint(0.5 * level.Value);

        ctx.Save(outputPath);

        var result = "List item edited successfully\n";
        result += $"Paragraph index: {paragraphIndex.Value}\n";
        result += $"New text: {text}\n";
        if (level.HasValue) result += $"Level: {level.Value}\n";
        result += ctx.GetOutputMessage(outputPath);

        return result;
    }

    /// <summary>
    ///     Sets list formatting options for a paragraph.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="paragraphIndex">The zero-based paragraph index.</param>
    /// <param name="numberStyle">The number style (arabic, roman, letter, bullet, none).</param>
    /// <param name="indentLevel">The indentation level (0-8).</param>
    /// <param name="leftIndent">The left indent in points.</param>
    /// <param name="firstLineIndent">The first line indent in points.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when paragraphIndex is not provided or out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when the paragraph cannot be found.</exception>
    private static string SetListFormat(DocumentContext<Document> ctx, string? outputPath, int? paragraphIndex,
        string? numberStyle, int? indentLevel, double? leftIndent, double? firstLineIndent)
    {
        if (!paragraphIndex.HasValue)
            throw new ArgumentException("paragraphIndex parameter is required for set_format operation");

        var doc = ctx.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
            throw new ArgumentException(
                $"Paragraph index {paragraphIndex.Value} is out of range (document has {paragraphs.Count} paragraphs)");

        var para = paragraphs[paragraphIndex.Value] as Paragraph;
        if (para == null)
            throw new InvalidOperationException($"Unable to find paragraph at index {paragraphIndex.Value}");

        List<string> changes = [];

        if (!string.IsNullOrEmpty(numberStyle) && para.ListFormat.IsListItem)
        {
            var list = para.ListFormat.List;
            if (list != null)
            {
                var level = para.ListFormat.ListLevelNumber;
                var listLevel = list.ListLevels[level];

                var style = numberStyle.ToLower() switch
                {
                    "arabic" => NumberStyle.Arabic,
                    "roman" => NumberStyle.UppercaseRoman,
                    "letter" => NumberStyle.UppercaseLetter,
                    "bullet" => NumberStyle.Bullet,
                    "none" => NumberStyle.None,
                    _ => NumberStyle.Arabic
                };

                listLevel.NumberStyle = style;
                changes.Add($"Number style: {numberStyle} (affects all items at level {level} in this list)");
            }
        }

        if (indentLevel.HasValue)
        {
            para.ParagraphFormat.LeftIndent = InchToPoint(0.5 * indentLevel.Value);
            changes.Add($"Indent level: {indentLevel.Value} ({InchToPoint(0.5 * indentLevel.Value):F1} points)");
        }

        if (leftIndent.HasValue)
        {
            para.ParagraphFormat.LeftIndent = leftIndent.Value;
            changes.Add($"Left indent: {leftIndent.Value} points");
        }

        if (firstLineIndent.HasValue)
        {
            para.ParagraphFormat.FirstLineIndent = firstLineIndent.Value;
            changes.Add($"First line indent: {firstLineIndent.Value} points");
        }

        ctx.Save(outputPath);

        var result = "List format set successfully\n";
        result += $"Paragraph index: {paragraphIndex.Value}\n";
        if (changes.Count > 0)
            result += $"Changes: {string.Join(", ", changes)}\n";
        else
            result += "No change parameters provided\n";
        result += ctx.GetOutputMessage(outputPath);

        return result;
    }

    /// <summary>
    ///     Gets list format information for a paragraph or all list paragraphs.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="paragraphIndex">The zero-based paragraph index, or null for all list paragraphs.</param>
    /// <returns>A JSON string containing list format information.</returns>
    /// <exception cref="ArgumentException">Thrown when the paragraph index is out of range.</exception>
    private static string GetListFormat(DocumentContext<Document> ctx, int? paragraphIndex)
    {
        var doc = ctx.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        // Build list item index mapping for each list
        var listItemIndices = new Dictionary<(int listId, int paraIndex), int>();
        var listCounters = new Dictionary<int, int>();
        foreach (var para in paragraphs)
            if (para.ListFormat is { IsListItem: true, List: not null })
            {
                var listId = para.ListFormat.List.ListId;
                listCounters.TryAdd(listId, 0);
                var paraIdx = paragraphs.IndexOf(para);
                listItemIndices[(listId, paraIdx)] = listCounters[listId];
                listCounters[listId]++;
            }

        if (paragraphIndex.HasValue)
        {
            if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
                throw new ArgumentException(
                    $"Paragraph index {paragraphIndex.Value} is out of range (document has {paragraphs.Count} paragraphs)");

            var para = paragraphs[paragraphIndex.Value];
            var listInfo = BuildListFormatInfo(para, paragraphIndex.Value, listItemIndices);

            return JsonSerializer.Serialize(listInfo, new JsonSerializerOptions { WriteIndented = true });
        }

        var listParagraphs = paragraphs
            .Where(p => p.ListFormat is { IsListItem: true })
            .ToList();

        if (listParagraphs.Count == 0)
        {
            var emptyResult = new
            {
                count = 0,
                listParagraphs = Array.Empty<object>(),
                message = "No list paragraphs found"
            };
            return JsonSerializer.Serialize(emptyResult, new JsonSerializerOptions { WriteIndented = true });
        }

        List<object> listInfos = [];
        foreach (var para in listParagraphs)
        {
            var paraIndex = paragraphs.IndexOf(para);
            listInfos.Add(BuildListFormatInfo(para, paraIndex, listItemIndices));
        }

        var result = new
        {
            count = listParagraphs.Count,
            listParagraphs = listInfos
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Builds list format information object for a paragraph.
    /// </summary>
    /// <param name="para">The paragraph to get list format information from.</param>
    /// <param name="paraIndex">The paragraph index.</param>
    /// <param name="listItemIndices">The dictionary mapping list items to their indices within lists.</param>
    /// <returns>An object containing the list format information.</returns>
    private static object BuildListFormatInfo(Paragraph para, int paraIndex,
        Dictionary<(int listId, int paraIndex), int> listItemIndices)
    {
        var previewText = para.ToString(SaveFormat.Text).Trim();
        if (previewText.Length > 50) previewText = previewText[..50] + "...";

        if (para.ListFormat is { IsListItem: true })
        {
            var listInfo = new Dictionary<string, object?>
            {
                ["paragraphIndex"] = paraIndex,
                ["contentPreview"] = previewText,
                ["isListItem"] = true,
                ["listLevel"] = para.ListFormat.ListLevelNumber
            };

            if (para.ListFormat.List != null)
            {
                var listId = para.ListFormat.List.ListId;
                listInfo["listId"] = listId;

                if (listItemIndices.TryGetValue((listId, paraIndex), out var listItemIndex))
                    listInfo["listItemIndex"] = listItemIndex;
            }

            if (para.ListFormat.ListLevel != null)
            {
                var level = para.ListFormat.ListLevel;
                listInfo["listLevelFormat"] = new
                {
                    symbol = level.NumberFormat,
                    alignment = level.Alignment.ToString(),
                    textPosition = level.TextPosition,
                    numberStyle = level.NumberStyle.ToString()
                };
            }

            return listInfo;
        }

        return new
        {
            paragraphIndex = paraIndex,
            contentPreview = previewText,
            isListItem = false,
            note = "This paragraph is not a list item. Use convert_to_list operation to convert it."
        };
    }

    /// <summary>
    ///     Restarts list numbering from a specified value at the given paragraph.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="paragraphIndex">The zero-based paragraph index.</param>
    /// <param name="startAt">The number to restart at.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when paragraphIndex is not provided, out of range, or paragraph is not a
    ///     list item.
    /// </exception>
    /// <exception cref="InvalidOperationException">Thrown when unable to access the list.</exception>
    private static string RestartNumbering(DocumentContext<Document> ctx, string? outputPath, int? paragraphIndex,
        int startAt)
    {
        if (!paragraphIndex.HasValue)
            throw new ArgumentException("paragraphIndex parameter is required for restart_numbering operation");

        var doc = ctx.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        if (paragraphIndex.Value < 0 || paragraphIndex.Value >= paragraphs.Count)
            throw new ArgumentException(
                $"Paragraph index {paragraphIndex.Value} is out of range (document has {paragraphs.Count} paragraphs)");

        var para = paragraphs[paragraphIndex.Value];

        if (!para.ListFormat.IsListItem)
            throw new ArgumentException(
                $"Paragraph at index {paragraphIndex.Value} is not a list item. Use get_format operation to find list item paragraphs.");

        var originalList = para.ListFormat.List;
        if (originalList == null)
            throw new InvalidOperationException("Unable to access list for this paragraph");

        // Create a copy of the list to restart numbering
        var newList = doc.Lists.AddCopy(originalList);
        var level = para.ListFormat.ListLevelNumber;

        // Set the starting number
        newList.ListLevels[level].StartAt = startAt;

        // Apply the new list to this paragraph and all following paragraphs in the same original list
        var applyCount = 0;
        for (var i = paragraphIndex.Value; i < paragraphs.Count; i++)
        {
            var p = paragraphs[i];
            if (p.ListFormat.IsListItem && p.ListFormat.List?.ListId == originalList.ListId)
            {
                p.ListFormat.List = newList;
                applyCount++;
            }
            else if (i > paragraphIndex.Value && !p.ListFormat.IsListItem)
            {
                break;
            }
        }

        ctx.Save(outputPath);

        var result = "List numbering restarted successfully\n";
        result += $"Paragraph index: {paragraphIndex.Value}\n";
        result += $"Start at: {startAt}\n";
        result += $"Paragraphs affected: {applyCount}\n";
        result += $"New list ID: {newList.ListId}\n";
        result += ctx.GetOutputMessage(outputPath);

        return result;
    }

    /// <summary>
    ///     Converts a range of paragraphs into a list.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="startIndex">The starting paragraph index.</param>
    /// <param name="endIndex">The ending paragraph index.</param>
    /// <param name="listType">The list type (bullet, number).</param>
    /// <param name="numberFormat">The number format (arabic, roman, letter).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when start or end index is not provided, out of range, or start is greater
    ///     than end.
    /// </exception>
    private static string ConvertToList(DocumentContext<Document> ctx, string? outputPath, int? startIndex,
        int? endIndex, string listType, string numberFormat)
    {
        if (!startIndex.HasValue)
            throw new ArgumentException("startParagraphIndex parameter is required for convert_to_list operation");
        if (!endIndex.HasValue)
            throw new ArgumentException("endParagraphIndex parameter is required for convert_to_list operation");

        var doc = ctx.Document;
        var paragraphs = doc.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        if (startIndex.Value < 0 || startIndex.Value >= paragraphs.Count)
            throw new ArgumentException(
                $"Start paragraph index {startIndex.Value} is out of range (document has {paragraphs.Count} paragraphs)");

        if (endIndex.Value < 0 || endIndex.Value >= paragraphs.Count)
            throw new ArgumentException(
                $"End paragraph index {endIndex.Value} is out of range (document has {paragraphs.Count} paragraphs)");

        if (startIndex.Value > endIndex.Value)
            throw new ArgumentException(
                $"Start index ({startIndex.Value}) must be less than or equal to end index ({endIndex.Value})");

        // Create list
        var list = doc.Lists.Add(listType == "number"
            ? ListTemplate.NumberDefault
            : ListTemplate.BulletDefault);

        // Configure number format if needed
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

        // Apply list to paragraphs
        var convertedCount = 0;
        var skippedCount = 0;
        for (var i = startIndex.Value; i <= endIndex.Value; i++)
        {
            var para = paragraphs[i];

            // Skip paragraphs that are already list items
            if (para.ListFormat.IsListItem)
            {
                skippedCount++;
                continue;
            }

            // Skip empty paragraphs
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

        ctx.Save(outputPath);

        var result = "Paragraphs converted to list successfully\n";
        result += $"Range: paragraph {startIndex.Value} to {endIndex.Value}\n";
        result += $"List type: {listType}\n";
        if (listType == "number") result += $"Number format: {numberFormat}\n";
        result += $"Converted: {convertedCount} paragraphs\n";
        if (skippedCount > 0) result += $"Skipped: {skippedCount} paragraphs (already list items or empty)\n";
        result += ctx.GetOutputMessage(outputPath);

        return result;
    }

    /// <summary>
    ///     Parses list items from JSON array supporting both string and object formats.
    /// </summary>
    /// <param name="itemsArray">The JSON array containing list items.</param>
    /// <returns>A list of tuples containing text and level for each item.</returns>
    /// <exception cref="ArgumentException">Thrown when items array is empty or contains invalid items.</exception>
    private static List<(string text, int level)> ParseItems(JsonArray itemsArray)
    {
        List<(string text, int level)> items = [];

        if (itemsArray.Count == 0)
            throw new ArgumentException("items array cannot be empty");

        foreach (var item in itemsArray)
        {
            if (item == null) continue;

            if (item is JsonValue jsonValue)
            {
                try
                {
                    var text = jsonValue.GetValue<string>();
                    if (!string.IsNullOrEmpty(text)) items.Add((text, 0));
                }
                catch (Exception ex)
                {
                    throw new ArgumentException($"Unable to parse list item as string: {ex.Message}");
                }
            }
            else if (item is JsonObject jsonObj)
            {
                var text = jsonObj["text"]?.GetValue<string>();
                if (string.IsNullOrEmpty(text))
                    throw new ArgumentException("List item object must contain 'text' property");

                var level = jsonObj["level"]?.GetValue<int>() ?? 0;
                if (level < 0 || level > 8) level = Math.Max(0, Math.Min(8, level));

                items.Add((text, level));
            }
            else
            {
                throw new ArgumentException($"Invalid list item format: {item.GetType().Name}");
            }
        }

        if (items.Count == 0)
            throw new ArgumentException("No valid list items after parsing");

        return items;
    }
}