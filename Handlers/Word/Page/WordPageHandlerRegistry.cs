using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Page;

/// <summary>
///     Registry for Word page operation handlers.
/// </summary>
public static class WordPageHandlerRegistry
{
    /// <summary>
    ///     Creates a handler registry with all page operation handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new SetMarginsWordHandler());
        registry.Register(new SetOrientationWordHandler());
        registry.Register(new SetSizeWordHandler());
        registry.Register(new SetPageNumberWordHandler());
        registry.Register(new SetPageSetupWordHandler());
        registry.Register(new DeletePageWordHandler());
        registry.Register(new InsertBlankPageWordHandler());
        registry.Register(new AddPageBreakWordHandler());
        return registry;
    }
}
