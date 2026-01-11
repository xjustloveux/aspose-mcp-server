using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Table;

/// <summary>
///     Registry for PowerPoint table operation handlers.
///     Provides a pre-configured registry with all table handlers registered.
/// </summary>
public static class PptTableHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PowerPoint table handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Presentation> Create()
    {
        var registry = new HandlerRegistry<Presentation>();
        registry.Register(new AddPptTableHandler());
        registry.Register(new EditPptTableHandler());
        registry.Register(new DeletePptTableHandler());
        registry.Register(new GetPptTableContentHandler());
        registry.Register(new InsertPptTableRowHandler());
        registry.Register(new InsertPptTableColumnHandler());
        registry.Register(new DeletePptTableRowHandler());
        registry.Register(new DeletePptTableColumnHandler());
        registry.Register(new EditPptTableCellHandler());
        return registry;
    }
}
