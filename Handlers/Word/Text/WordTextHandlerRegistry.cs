using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Text;

/// <summary>
///     Registry for Word text operation handlers.
/// </summary>
public static class WordTextHandlerRegistry
{
    /// <summary>
    ///     Creates and populates a handler registry with all Word text handlers.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new AddWordTextHandler());
        registry.Register(new DeleteWordTextHandler());
        registry.Register(new ReplaceWordTextHandler());
        registry.Register(new SearchWordTextHandler());
        registry.Register(new FormatWordTextHandler());
        registry.Register(new InsertAtPositionWordTextHandler());
        registry.Register(new DeleteRangeWordTextHandler());
        registry.Register(new AddWithStyleWordTextHandler());
        return registry;
    }
}
