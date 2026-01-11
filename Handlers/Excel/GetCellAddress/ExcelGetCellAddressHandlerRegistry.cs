using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.GetCellAddress;

/// <summary>
///     Registry for Excel get cell address operation handlers.
///     Provides a pre-configured registry with all cell address handlers registered.
/// </summary>
public static class ExcelGetCellAddressHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Excel get cell address handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Workbook> Create()
    {
        var registry = new HandlerRegistry<Workbook>();
        registry.Register(new FromA1NotationHandler());
        registry.Register(new FromIndexHandler());
        return registry;
    }
}
