using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.Protect;

/// <summary>
///     Registry for Excel protect operation handlers.
///     Provides a pre-configured registry with all protect handlers registered.
/// </summary>
public static class ExcelProtectHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Excel protect handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Workbook> Create()
    {
        var registry = new HandlerRegistry<Workbook>();
        registry.Register(new ProtectExcelHandler());
        registry.Register(new UnprotectExcelHandler());
        registry.Register(new GetExcelProtectionHandler());
        registry.Register(new SetExcelCellLockedHandler());
        return registry;
    }
}
