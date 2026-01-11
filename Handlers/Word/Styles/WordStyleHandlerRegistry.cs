using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Styles;

/// <summary>
///     Registry for Word style operation handlers.
///     Provides a pre-configured registry with all style handlers registered.
/// </summary>
public static class WordStyleHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Word style handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new GetWordStylesHandler());
        registry.Register(new CreateWordStyleHandler());
        registry.Register(new ApplyWordStyleHandler());
        registry.Register(new CopyWordStylesHandler());
        return registry;
    }
}
