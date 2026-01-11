using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Note;

/// <summary>
///     Registry for Word note operation handlers.
///     Provides a pre-configured registry with all footnote and endnote handlers registered.
/// </summary>
public static class WordNoteHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Word note handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new AddWordFootnoteHandler());
        registry.Register(new AddWordEndnoteHandler());
        registry.Register(new DeleteWordFootnoteHandler());
        registry.Register(new DeleteWordEndnoteHandler());
        registry.Register(new GetWordFootnotesHandler());
        registry.Register(new GetWordEndnotesHandler());
        registry.Register(new EditWordFootnoteHandler());
        registry.Register(new EditWordEndnoteHandler());
        return registry;
    }
}
