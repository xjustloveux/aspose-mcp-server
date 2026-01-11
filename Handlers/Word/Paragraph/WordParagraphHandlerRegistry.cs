using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Paragraph;

/// <summary>
///     Registry for Word paragraph operation handlers.
/// </summary>
public static class WordParagraphHandlerRegistry
{
    /// <summary>
    ///     Creates a handler registry with all paragraph operation handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new InsertParagraphWordHandler());
        registry.Register(new DeleteParagraphWordHandler());
        registry.Register(new EditParagraphWordHandler());
        registry.Register(new GetParagraphsWordHandler());
        registry.Register(new GetParagraphFormatWordHandler());
        registry.Register(new CopyParagraphFormatWordHandler());
        registry.Register(new MergeParagraphsWordHandler());
        return registry;
    }
}
