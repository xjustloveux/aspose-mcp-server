using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Image;

/// <summary>
///     Registry for Word image operation handlers.
/// </summary>
public static class WordImageHandlerRegistry
{
    /// <summary>
    ///     Creates a handler registry with all image operation handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new AddImageWordHandler());
        registry.Register(new EditImageWordHandler());
        registry.Register(new DeleteImageWordHandler());
        registry.Register(new GetImagesWordHandler());
        registry.Register(new ReplaceImageWordHandler());
        registry.Register(new ExtractImagesWordHandler());
        return registry;
    }
}
