using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Attachment;

/// <summary>
///     Registry for PDF attachment operation handlers.
/// </summary>
public static class PdfAttachmentHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PDF attachment handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new AddPdfAttachmentHandler());
        registry.Register(new DeletePdfAttachmentHandler());
        registry.Register(new GetPdfAttachmentsHandler());
        return registry;
    }
}
