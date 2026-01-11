using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Hyperlink;

/// <summary>
///     Registry for Word hyperlink operation handlers.
/// </summary>
public static class WordHyperlinkHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Word hyperlink handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new AddWordHyperlinkHandler());
        registry.Register(new EditWordHyperlinkHandler());
        registry.Register(new DeleteWordHyperlinkHandler());
        registry.Register(new GetWordHyperlinksHandler());
        return registry;
    }
}
