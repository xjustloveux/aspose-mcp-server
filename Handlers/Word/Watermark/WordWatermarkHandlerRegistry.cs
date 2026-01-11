using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Watermark;

/// <summary>
///     Registry for Word watermark operation handlers.
/// </summary>
public static class WordWatermarkHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Word watermark handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new AddTextWatermarkWordHandler());
        registry.Register(new AddImageWatermarkWordHandler());
        registry.Register(new RemoveWatermarkWordHandler());
        return registry;
    }
}
