using System.Text.Json.Nodes;
using Aspose.Pdf;

namespace AsposeMcpServer.Tools;

public class PdfDeleteBookmarkTool : IAsposeTool
{
    public string Description => "Delete bookmark(s) from PDF document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "PDF file path"
            },
            bookmarkIndex = new
            {
                type = "number",
                description = "Bookmark index to delete (0-based)"
            },
            bookmarkIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Array of bookmark indices to delete (0-based, optional, overrides bookmarkIndex)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var bookmarkIndex = arguments?["bookmarkIndex"]?.GetValue<int?>();
        var bookmarkIndicesArray = arguments?["bookmarkIndices"]?.AsArray();

        using var document = new Document(path);
        var outlines = document.Outlines;

        List<int> bookmarksToDelete;
        if (bookmarkIndicesArray != null && bookmarkIndicesArray.Count > 0)
        {
            bookmarksToDelete = bookmarkIndicesArray.Select(b => b?.GetValue<int>()).Where(b => b.HasValue).Select(b => b!.Value).OrderByDescending(b => b).ToList();
        }
        else if (bookmarkIndex.HasValue)
        {
            bookmarksToDelete = new List<int> { bookmarkIndex.Value };
        }
        else
        {
            throw new ArgumentException("Either bookmarkIndex or bookmarkIndices must be provided");
        }

        // Delete bookmarks in reverse order to maintain indices
        var sortedIndices = bookmarksToDelete.OrderByDescending(i => i).ToList();
        foreach (var index in sortedIndices)
        {
            if (index < 0 || index >= outlines.Count)
            {
                continue;
            }
            // Remove bookmark by getting the bookmark object first
            var bookmark = outlines[index];
            outlines.Remove(bookmark);
        }

        document.Save(path);
        return await Task.FromResult($"Deleted {bookmarksToDelete.Count} bookmark(s). Remaining bookmarks: {outlines.Count}");
    }
}

