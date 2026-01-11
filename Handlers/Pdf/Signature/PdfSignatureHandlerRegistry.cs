using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Signature;

/// <summary>
///     Registry for PDF signature operation handlers.
/// </summary>
public static class PdfSignatureHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PDF signature handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new SignPdfHandler());
        registry.Register(new DeletePdfSignatureHandler());
        registry.Register(new GetPdfSignaturesHandler());
        return registry;
    }
}
