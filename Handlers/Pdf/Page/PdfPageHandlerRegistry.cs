using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Page;

/// <summary>
///     Registry for PDF page operation handlers.
///     Provides a pre-configured registry with all page handlers registered.
/// </summary>
public static class PdfPageHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PDF page handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new AddPdfPageHandler());
        registry.Register(new DeletePdfPageHandler());
        registry.Register(new RotatePdfPageHandler());
        registry.Register(new GetPdfPageDetailsHandler());
        registry.Register(new GetPdfPageInfoHandler());
        return registry;
    }
}
