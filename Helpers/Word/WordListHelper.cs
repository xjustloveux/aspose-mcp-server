using System.Text.Json.Nodes;
using Aspose.Words;
using AsposeMcpServer.Results.Word.List;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Helpers.Word;

/// <summary>
///     Helper class for Word list operations.
/// </summary>
public static class WordListHelper
{
    /// <summary>
    ///     Parses list items from JSON array supporting both string and object formats.
    /// </summary>
    /// <param name="itemsArray">The JSON array containing list items.</param>
    /// <returns>A list of <see cref="ListItemInfo" /> for each item.</returns>
    /// <exception cref="ArgumentException">Thrown when items array is empty or contains invalid items.</exception>
    public static List<ListItemInfo> ParseItems(JsonArray itemsArray)
    {
        if (itemsArray.Count == 0)
            throw new ArgumentException("items array cannot be empty");

        List<ListItemInfo> items = [];

        foreach (var item in itemsArray)
        {
            if (item == null) continue;

            var parsed = ParseSingleItem(item);
            if (parsed != null)
                items.Add(parsed);
        }

        if (items.Count == 0)
            throw new ArgumentException("No valid list items after parsing");

        return items;
    }

    /// <summary>
    ///     Parses a single list item from a JSON node.
    /// </summary>
    /// <param name="item">The JSON node to parse.</param>
    /// <returns>A <see cref="ListItemInfo" />, or null if invalid.</returns>
    /// <exception cref="ArgumentException">Thrown when item format is invalid.</exception>
    private static ListItemInfo? ParseSingleItem(JsonNode item)
    {
        return item switch
        {
            JsonValue jsonValue => ParseJsonValueItem(jsonValue),
            JsonObject jsonObj => ParseJsonObjectItem(jsonObj),
            _ => throw new ArgumentException($"Invalid list item format: {item.GetType().Name}")
        };
    }

    /// <summary>
    ///     Parses a list item from a JSON value (string).
    /// </summary>
    /// <param name="jsonValue">The JSON value to parse.</param>
    /// <returns>A <see cref="ListItemInfo" /> with level 0, or null if empty.</returns>
    /// <exception cref="ArgumentException">Thrown when value cannot be parsed.</exception>
    private static ListItemInfo? ParseJsonValueItem(JsonValue jsonValue)
    {
        try
        {
            var text = jsonValue.GetValue<string>();
            return string.IsNullOrEmpty(text) ? null : new ListItemInfo(text, 0);
        }
        catch (Exception ex)
        {
            throw new ArgumentException($"Unable to parse list item as string: {ex.Message}");
        }
    }

    /// <summary>
    ///     Parses a list item from a JSON object with text and level properties.
    /// </summary>
    /// <param name="jsonObj">The JSON object to parse.</param>
    /// <returns>A <see cref="ListItemInfo" /> with text and level.</returns>
    /// <exception cref="ArgumentException">Thrown when text property is missing.</exception>
    private static ListItemInfo ParseJsonObjectItem(JsonObject jsonObj)
    {
        var text = jsonObj["text"]?.GetValue<string>();
        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("List item object must contain 'text' property");

        var level = jsonObj["level"]?.GetValue<int>() ?? 0;
        level = Math.Max(0, Math.Min(8, level));

        return new ListItemInfo(text, level);
    }

    /// <summary>
    ///     Builds list format information object for a paragraph.
    /// </summary>
    /// <param name="para">The paragraph to get list format information from.</param>
    /// <param name="paraIndex">The paragraph index.</param>
    /// <param name="listItemIndices">The dictionary mapping list items to their indices within lists.</param>
    /// <returns>An object containing the list format information.</returns>
    public static object BuildListFormatInfo(WordParagraph para, int paraIndex,
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
    ///     Builds list format single result for a paragraph.
    /// </summary>
    /// <param name="para">The paragraph to get list format information from.</param>
    /// <param name="paraIndex">The paragraph index.</param>
    /// <param name="listItemIndices">The dictionary mapping list items to their indices within lists.</param>
    /// <returns>A GetWordListFormatSingleResult containing the list format information.</returns>
    public static GetWordListFormatSingleResult BuildListFormatSingleResult(WordParagraph para, int paraIndex,
        Dictionary<(int listId, int paraIndex), int> listItemIndices)
    {
        var previewText = para.ToString(SaveFormat.Text).Trim();
        if (previewText.Length > 50) previewText = previewText[..50] + "...";

        if (para.ListFormat is { IsListItem: true })
        {
            int? listId = null;
            int? listItemIndex = null;

            if (para.ListFormat.List != null)
            {
                listId = para.ListFormat.List.ListId;
                if (listItemIndices.TryGetValue((listId.Value, paraIndex), out var idx))
                    listItemIndex = idx;
            }

            ListLevelFormatInfo? levelFormat = null;
            if (para.ListFormat.ListLevel != null)
            {
                var level = para.ListFormat.ListLevel;
                levelFormat = new ListLevelFormatInfo
                {
                    Symbol = level.NumberFormat,
                    Alignment = level.Alignment.ToString(),
                    TextPosition = level.TextPosition,
                    NumberStyle = level.NumberStyle.ToString()
                };
            }

            return new GetWordListFormatSingleResult
            {
                ParagraphIndex = paraIndex,
                ContentPreview = previewText,
                IsListItem = true,
                ListLevel = para.ListFormat.ListLevelNumber,
                ListId = listId,
                ListItemIndex = listItemIndex,
                ListLevelFormat = levelFormat
            };
        }

        return new GetWordListFormatSingleResult
        {
            ParagraphIndex = paraIndex,
            ContentPreview = previewText,
            IsListItem = false,
            Note = "This paragraph is not a list item. Use convert_to_list operation to convert it."
        };
    }

    /// <summary>
    ///     Builds list paragraph info for a paragraph that is a list item.
    /// </summary>
    /// <param name="para">The paragraph to get list format information from.</param>
    /// <param name="paraIndex">The paragraph index.</param>
    /// <param name="listItemIndices">The dictionary mapping list items to their indices within lists.</param>
    /// <returns>A ListParagraphInfo containing the list format information.</returns>
    public static ListParagraphInfo BuildListParagraphInfo(WordParagraph para, int paraIndex,
        Dictionary<(int listId, int paraIndex), int> listItemIndices)
    {
        var previewText = para.ToString(SaveFormat.Text).Trim();
        if (previewText.Length > 50) previewText = previewText[..50] + "...";

        int? listId = null;
        int? listItemIndex = null;

        if (para.ListFormat?.List != null)
        {
            listId = para.ListFormat.List.ListId;
            if (listItemIndices.TryGetValue((listId.Value, paraIndex), out var idx))
                listItemIndex = idx;
        }

        ListLevelFormatInfo? levelFormat = null;
        if (para.ListFormat?.ListLevel != null)
        {
            var level = para.ListFormat.ListLevel;
            levelFormat = new ListLevelFormatInfo
            {
                Symbol = level.NumberFormat,
                Alignment = level.Alignment.ToString(),
                TextPosition = level.TextPosition,
                NumberStyle = level.NumberStyle.ToString()
            };
        }

        return new ListParagraphInfo
        {
            ParagraphIndex = paraIndex,
            ContentPreview = previewText,
            IsListItem = para.ListFormat is { IsListItem: true },
            ListLevel = para.ListFormat?.ListLevelNumber,
            ListId = listId,
            ListItemIndex = listItemIndex,
            ListLevelFormat = levelFormat
        };
    }
}
