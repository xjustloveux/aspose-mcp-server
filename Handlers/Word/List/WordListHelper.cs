using System.Text.Json.Nodes;
using Aspose.Words;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.List;

/// <summary>
///     Helper class for Word list operations.
/// </summary>
public static class WordListHelper
{
    /// <summary>
    ///     Parses list items from JSON array supporting both string and object formats.
    /// </summary>
    /// <param name="itemsArray">The JSON array containing list items.</param>
    /// <returns>A list of tuples containing text and level for each item.</returns>
    /// <exception cref="ArgumentException">Thrown when items array is empty or contains invalid items.</exception>
    public static List<(string text, int level)> ParseItems(JsonArray itemsArray)
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
}
