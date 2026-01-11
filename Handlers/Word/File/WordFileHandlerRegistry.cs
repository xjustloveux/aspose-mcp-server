using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.File;

/// <summary>
///     Registry for Word file operation handlers.
/// </summary>
public static class WordFileHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Word file handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new CreateWordDocumentHandler());
        registry.Register(new CreateFromTemplateWordHandler());
        registry.Register(new ConvertWordDocumentHandler());
        registry.Register(new MergeWordDocumentsHandler());
        registry.Register(new SplitWordDocumentHandler());
        return registry;
    }
}
