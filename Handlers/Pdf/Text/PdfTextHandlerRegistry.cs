using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Text;

/// <summary>
///     Registry for PDF text operation handlers.
///     Provides a pre-configured registry with all text handlers registered.
/// </summary>
public static class PdfTextHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PDF text handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new AddPdfTextHandler());
        registry.Register(new EditPdfTextHandler());
        registry.Register(new ExtractPdfTextHandler());
        return registry;
    }
}
