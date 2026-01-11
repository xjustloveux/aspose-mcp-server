using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Annotation;

/// <summary>
///     Registry for PDF annotation operation handlers.
///     Provides a pre-configured registry with all annotation handlers registered.
/// </summary>
public static class PdfAnnotationHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PDF annotation handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new AddPdfAnnotationHandler());
        registry.Register(new DeletePdfAnnotationHandler());
        registry.Register(new EditPdfAnnotationHandler());
        registry.Register(new GetPdfAnnotationsHandler());
        return registry;
    }
}
