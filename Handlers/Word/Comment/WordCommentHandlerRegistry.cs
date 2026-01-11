using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Comment;

/// <summary>
///     Registry for Word comment operation handlers.
/// </summary>
public static class WordCommentHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Word comment handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new AddWordCommentHandler());
        registry.Register(new DeleteWordCommentHandler());
        registry.Register(new GetWordCommentsHandler());
        registry.Register(new ReplyWordCommentHandler());
        return registry;
    }
}
