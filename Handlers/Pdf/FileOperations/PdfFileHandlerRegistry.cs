using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.FileOperations;

/// <summary>
///     Registry for PDF file operation handlers.
///     Provides a pre-configured registry with all file handlers registered.
/// </summary>
public static class PdfFileHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PDF file handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new CreatePdfFileHandler());
        registry.Register(new MergePdfFilesHandler());
        registry.Register(new SplitPdfFileHandler());
        registry.Register(new CompressPdfFileHandler());
        registry.Register(new EncryptPdfFileHandler());
        registry.Register(new LinearizePdfFileHandler());
        return registry;
    }
}
