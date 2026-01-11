using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Shape;

/// <summary>
///     Registry for Word shape operation handlers.
/// </summary>
public static class WordShapeHandlerRegistry
{
    /// <summary>
    ///     Creates a handler registry with all shape operation handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new AddLineWordHandler());
        registry.Register(new AddTextBoxWordHandler());
        registry.Register(new GetTextboxesWordHandler());
        registry.Register(new EditTextBoxContentWordHandler());
        registry.Register(new SetTextBoxBorderWordHandler());
        registry.Register(new AddChartWordHandler());
        registry.Register(new AddShapeWordHandler());
        registry.Register(new GetShapesWordHandler());
        registry.Register(new DeleteShapeWordHandler());
        return registry;
    }
}
