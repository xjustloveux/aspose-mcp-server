using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Content;

/// <summary>
///     Registry for Word content operation handlers.
///     Provides a pre-configured registry with all content handlers registered.
/// </summary>
public static class WordContentHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Word content handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new GetWordContentHandler());
        registry.Register(new GetWordContentDetailedHandler());
        registry.Register(new GetWordStatisticsHandler());
        registry.Register(new GetWordDocumentInfoHandler());
        return registry;
    }
}
