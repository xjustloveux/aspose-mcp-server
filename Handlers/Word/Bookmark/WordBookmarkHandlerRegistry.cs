using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Bookmark;

/// <summary>
///     Registry for Word bookmark operation handlers.
/// </summary>
public static class WordBookmarkHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Word bookmark handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new AddWordBookmarkHandler());
        registry.Register(new EditWordBookmarkHandler());
        registry.Register(new DeleteWordBookmarkHandler());
        registry.Register(new GetWordBookmarksHandler());
        registry.Register(new GotoWordBookmarkHandler());
        return registry;
    }
}
