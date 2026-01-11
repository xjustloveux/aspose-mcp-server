using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Image;

/// <summary>
///     Registry for PDF image operation handlers.
///     Provides a pre-configured registry with all image handlers registered.
/// </summary>
public static class PdfImageHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PDF image handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new AddPdfImageHandler());
        registry.Register(new DeletePdfImageHandler());
        registry.Register(new EditPdfImageHandler());
        registry.Register(new ExtractPdfImageHandler());
        registry.Register(new GetPdfImagesHandler());
        return registry;
    }
}
