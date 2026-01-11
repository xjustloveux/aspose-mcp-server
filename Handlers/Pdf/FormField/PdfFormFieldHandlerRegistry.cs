using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.FormField;

/// <summary>
///     Registry for PDF form field operation handlers.
/// </summary>
public static class PdfFormFieldHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PDF form field handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new AddPdfFormFieldHandler());
        registry.Register(new DeletePdfFormFieldHandler());
        registry.Register(new EditPdfFormFieldHandler());
        registry.Register(new GetPdfFormFieldsHandler());
        return registry;
    }
}
