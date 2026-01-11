using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.Group;

/// <summary>
///     Registry for Excel group operation handlers.
///     Provides a pre-configured registry with all group handlers registered.
/// </summary>
public static class ExcelGroupHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Excel group handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Workbook> Create()
    {
        var registry = new HandlerRegistry<Workbook>();
        registry.Register(new GroupExcelRowsHandler());
        registry.Register(new UngroupExcelRowsHandler());
        registry.Register(new GroupExcelColumnsHandler());
        registry.Register(new UngroupExcelColumnsHandler());
        return registry;
    }
}
