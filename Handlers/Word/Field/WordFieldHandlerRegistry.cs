using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Field;

/// <summary>
///     Registry for Word field operation handlers.
/// </summary>
public static class WordFieldHandlerRegistry
{
    /// <summary>
    ///     Creates a handler registry with all field operation handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new InsertFieldWordHandler());
        registry.Register(new EditFieldWordHandler());
        registry.Register(new DeleteFieldWordHandler());
        registry.Register(new UpdateFieldWordHandler());
        registry.Register(new GetFieldsWordHandler());
        registry.Register(new GetFieldDetailWordHandler());
        registry.Register(new AddFormFieldWordHandler());
        registry.Register(new EditFormFieldWordHandler());
        registry.Register(new DeleteFormFieldWordHandler());
        registry.Register(new GetFormFieldsWordHandler());
        return registry;
    }
}
