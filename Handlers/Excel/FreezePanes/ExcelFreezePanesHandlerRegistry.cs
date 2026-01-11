using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.FreezePanes;

/// <summary>
///     Registry for Excel freeze panes operation handlers.
///     Provides a pre-configured registry with all freeze panes handlers registered.
/// </summary>
public static class ExcelFreezePanesHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Excel freeze panes handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Workbook> Create()
    {
        var registry = new HandlerRegistry<Workbook>();
        registry.Register(new FreezeExcelPanesHandler());
        registry.Register(new UnfreezeExcelPanesHandler());
        registry.Register(new GetExcelFreezePanesHandler());
        return registry;
    }
}
