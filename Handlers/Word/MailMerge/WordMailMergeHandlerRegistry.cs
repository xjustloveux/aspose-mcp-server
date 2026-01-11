using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.MailMerge;

/// <summary>
///     Registry for Word mail merge operation handlers.
/// </summary>
public static class WordMailMergeHandlerRegistry
{
    /// <summary>
    ///     Creates and populates a handler registry with all Word mail merge handlers.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new ExecuteMailMergeHandler());
        return registry;
    }
}
