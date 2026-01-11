using System.Text.Json;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Bookmark;

/// <summary>
///     Handler for getting bookmarks from Word documents.
/// </summary>
public class GetWordBookmarksHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all bookmarks from the document as JSON.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">No parameters required.</param>
    /// <returns>A JSON string containing all bookmarks information.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var doc = context.Document;
        var bookmarks = doc.Range.Bookmarks;

        if (bookmarks.Count == 0)
            return JsonSerializer.Serialize(new
                { count = 0, bookmarks = Array.Empty<object>(), message = "No bookmarks found in document" });

        List<object> bookmarkList = [];
        var index = 0;
        foreach (var bookmark in bookmarks)
        {
            bookmarkList.Add(new
            {
                index,
                name = bookmark.Name,
                text = bookmark.Text,
                length = bookmark.Text.Length
            });
            index++;
        }

        var result = new
        {
            count = bookmarks.Count,
            bookmarks = bookmarkList
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }
}
