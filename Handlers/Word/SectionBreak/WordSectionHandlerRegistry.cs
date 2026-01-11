using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.SectionBreak;

/// <summary>
///     Registry for Word section operation handlers.
///     Provides a pre-configured registry with all section handlers registered.
/// </summary>
public static class WordSectionHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Word section handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new InsertWordSectionHandler());
        registry.Register(new DeleteWordSectionHandler());
        registry.Register(new GetWordSectionsHandler());
        return registry;
    }
}
