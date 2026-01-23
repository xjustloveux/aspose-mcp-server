using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Pdf.Bookmark;

namespace AsposeMcpServer.Handlers.Pdf.Bookmark;

/// <summary>
///     Handler for getting bookmarks from PDF documents.
/// </summary>
[ResultType(typeof(GetBookmarksPdfResult))]
public class GetPdfBookmarksHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Retrieves all bookmarks from the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">No required parameters.</param>
    /// <returns>JSON string containing bookmark information.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var document = context.Document;

        if (document.Outlines.Count == 0)
            return new GetBookmarksPdfResult
            {
                Count = 0,
                Items = [],
                Message = "No bookmarks found"
            };

        var bookmarkList = new List<PdfBookmarkInfo>();
        var index = 0;

        foreach (var bookmark in document.Outlines)
        {
            index++;
            var extractedPageIndex = ExtractPageIndex(bookmark, document);

            bookmarkList.Add(new PdfBookmarkInfo
            {
                Index = index,
                Title = bookmark.Title,
                PageIndex = extractedPageIndex
            });
        }

        return new GetBookmarksPdfResult
        {
            Count = bookmarkList.Count,
            Items = bookmarkList
        };
    }

    /// <summary>
    ///     Extracts the target page index from a bookmark.
    /// </summary>
    /// <param name="bookmark">The bookmark to extract the page index from.</param>
    /// <param name="document">The PDF document.</param>
    /// <returns>The 1-based page index, or null if not available.</returns>
    private static int? ExtractPageIndex(OutlineItemCollection bookmark, Document document)
    {
        Aspose.Pdf.Page? page = null;

        if (bookmark.Destination is ExplicitDestination explicitDest)
            page = explicitDest.Page;
        else if (bookmark.Action is GoToAction { Destination: ExplicitDestination actionDest })
            page = actionDest.Page;

        if (page == null)
            return null;

        var pageIndex = document.Pages.IndexOf(page);
        return pageIndex > 0 ? pageIndex : null;
    }
}
