using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Properties;

/// <summary>
///     Registry for PDF properties operation handlers.
/// </summary>
public static class PdfPropertiesHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PDF properties handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new GetPdfPropertiesHandler());
        registry.Register(new SetPdfPropertiesHandler());
        return registry;
    }
}
