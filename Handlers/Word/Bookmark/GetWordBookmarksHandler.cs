using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Word.Bookmark;

namespace AsposeMcpServer.Handlers.Word.Bookmark;

/// <summary>
///     Handler for getting bookmarks from Word documents.
/// </summary>
[ResultType(typeof(GetBookmarksResult))]
public class GetWordBookmarksHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all bookmarks from the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">No parameters required.</param>
    /// <returns>Result containing all bookmarks information.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        _ = parameters;

        var doc = context.Document;
        var bookmarks = doc.Range.Bookmarks;

        if (bookmarks.Count == 0)
            return new GetBookmarksResult
            {
                Count = 0,
                Bookmarks = [],
                Message = "No bookmarks found in document"
            };

        var bookmarkList = new List<BookmarkInfo>();
        var index = 0;
        foreach (var bookmark in bookmarks)
        {
            bookmarkList.Add(new BookmarkInfo
            {
                Index = index,
                Name = bookmark.Name,
                Text = bookmark.Text,
                Length = bookmark.Text.Length
            });
            index++;
        }

        return new GetBookmarksResult
        {
            Count = bookmarks.Count,
            Bookmarks = bookmarkList
        };
    }
}
