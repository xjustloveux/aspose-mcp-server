using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Properties;

/// <summary>
///     Registry for Word properties operation handlers.
/// </summary>
public static class WordPropertiesHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Word properties handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new GetWordPropertiesHandler());
        registry.Register(new SetWordPropertiesHandler());
        return registry;
    }
}
