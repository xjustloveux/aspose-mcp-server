using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Protection;

/// <summary>
///     Registry for Word protection operation handlers.
/// </summary>
public static class WordProtectionHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Word protection handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new ProtectWordHandler());
        registry.Register(new UnprotectWordHandler());
        return registry;
    }
}
