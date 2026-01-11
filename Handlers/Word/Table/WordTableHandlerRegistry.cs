using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Table;

/// <summary>
///     Registry for Word table operation handlers.
/// </summary>
public static class WordTableHandlerRegistry
{
    /// <summary>
    ///     Creates and populates a handler registry with all Word table handlers.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Document> Create()
    {
        var registry = new HandlerRegistry<Document>();
        registry.Register(new CreateWordTableHandler());
        registry.Register(new DeleteWordTableHandler());
        registry.Register(new GetWordTablesHandler());
        registry.Register(new InsertRowWordTableHandler());
        registry.Register(new DeleteRowWordTableHandler());
        registry.Register(new InsertColumnWordTableHandler());
        registry.Register(new DeleteColumnWordTableHandler());
        registry.Register(new MergeCellsWordTableHandler());
        registry.Register(new SplitCellWordTableHandler());
        registry.Register(new EditCellFormatWordTableHandler());
        registry.Register(new MoveWordTableHandler());
        registry.Register(new CopyWordTableHandler());
        registry.Register(new GetStructureWordTableHandler());
        registry.Register(new SetBorderWordTableHandler());
        registry.Register(new SetColumnWidthWordTableHandler());
        registry.Register(new SetRowHeightWordTableHandler());
        return registry;
    }
}
