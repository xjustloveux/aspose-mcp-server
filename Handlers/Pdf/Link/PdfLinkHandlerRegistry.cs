using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Link;

/// <summary>
///     Registry for PDF link operation handlers.
/// </summary>
public static class PdfLinkHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PDF link handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new AddPdfLinkHandler());
        registry.Register(new DeletePdfLinkHandler());
        registry.Register(new EditPdfLinkHandler());
        registry.Register(new GetPdfLinksHandler());
        return registry;
    }
}
