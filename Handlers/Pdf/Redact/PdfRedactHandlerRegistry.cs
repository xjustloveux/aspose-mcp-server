using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Redact;

/// <summary>
///     Registry for PDF redaction operation handlers.
/// </summary>
public static class PdfRedactHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PDF redaction handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new RedactAreaHandler());
        registry.Register(new RedactTextHandler());
        return registry;
    }
}
