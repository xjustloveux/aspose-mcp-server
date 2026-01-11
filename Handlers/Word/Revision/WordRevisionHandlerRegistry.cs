using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Revision;

/// <summary>
///     Registry for Word revision operation handlers.
///     Provides a pre-configured registry with all revision handlers registered.
/// </summary>
public static class WordRevisionHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Word revision handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new GetRevisionsHandler());
        registry.Register(new AcceptAllRevisionsHandler());
        registry.Register(new RejectAllRevisionsHandler());
        registry.Register(new ManageRevisionHandler());
        registry.Register(new CompareDocumentsHandler());
        return registry;
    }
}
