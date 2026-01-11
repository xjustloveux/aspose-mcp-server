using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Table;

/// <summary>
///     Registry for PDF table operation handlers.
///     Provides a pre-configured registry with all table handlers registered.
/// </summary>
public static class PdfTableHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PDF table handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new AddPdfTableHandler());
        registry.Register(new EditPdfTableHandler());
        return registry;
    }
}
