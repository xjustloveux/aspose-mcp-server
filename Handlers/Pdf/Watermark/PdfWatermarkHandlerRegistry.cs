using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Watermark;

/// <summary>
///     Registry for PDF watermark operation handlers.
///     Provides a pre-configured registry with all watermark handlers registered.
/// </summary>
public static class PdfWatermarkHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PDF watermark handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new AddPdfWatermarkHandler());
        return registry;
    }
}
