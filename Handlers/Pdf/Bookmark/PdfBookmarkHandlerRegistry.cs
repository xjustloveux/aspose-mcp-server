using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Bookmark;

/// <summary>
///     Registry for PDF bookmark operation handlers.
///     Provides a pre-configured registry with all bookmark handlers registered.
/// </summary>
public static class PdfBookmarkHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PDF bookmark handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new AddPdfBookmarkHandler());
        registry.Register(new DeletePdfBookmarkHandler());
        registry.Register(new EditPdfBookmarkHandler());
        registry.Register(new GetPdfBookmarksHandler());
        return registry;
    }
}
