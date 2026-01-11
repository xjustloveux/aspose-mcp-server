using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.List;

/// <summary>
///     Registry for Word list operation handlers.
///     Provides a pre-configured registry with all list handlers registered.
/// </summary>
public static class WordListHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Word list handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new AddWordListHandler());
        registry.Register(new AddWordListItemHandler());
        registry.Register(new DeleteWordListItemHandler());
        registry.Register(new EditWordListItemHandler());
        registry.Register(new SetWordListFormatHandler());
        registry.Register(new GetWordListFormatHandler());
        registry.Register(new RestartWordListNumberingHandler());
        registry.Register(new ConvertToWordListHandler());
        return registry;
    }
}
