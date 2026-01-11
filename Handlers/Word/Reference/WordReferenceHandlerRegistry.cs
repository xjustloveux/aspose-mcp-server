using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Reference;

/// <summary>
///     Registry for Word reference operation handlers.
/// </summary>
public static class WordReferenceHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Word reference handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new AddTableOfContentsWordHandler());
        registry.Register(new UpdateTableOfContentsWordHandler());
        registry.Register(new AddIndexWordHandler());
        registry.Register(new AddCrossReferenceWordHandler());
        return registry;
    }
}
