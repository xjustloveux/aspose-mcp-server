using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Info;

/// <summary>
///     Registry for PDF info operation handlers.
/// </summary>
public static class PdfInfoHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PDF info handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new GetPdfContentHandler());
        registry.Register(new GetPdfStatisticsHandler());
        return registry;
    }
}
