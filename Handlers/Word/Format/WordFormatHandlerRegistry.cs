using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Format;

/// <summary>
///     Registry for Word format operation handlers.
/// </summary>
public static class WordFormatHandlerRegistry
{
    /// <summary>
    ///     Creates a handler registry with all format operation handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new GetRunFormatWordHandler());
        registry.Register(new SetRunFormatWordHandler());
        registry.Register(new GetTabStopsWordHandler());
        registry.Register(new AddTabStopWordHandler());
        registry.Register(new ClearTabStopsWordHandler());
        registry.Register(new SetParagraphBorderWordHandler());
        return registry;
    }
}
