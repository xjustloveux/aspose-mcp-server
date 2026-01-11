using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.Cell;

/// <summary>
///     Registry for Excel cell operation handlers.
///     Provides a pre-configured registry with all cell handlers registered.
/// </summary>
public static class ExcelCellHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Excel cell handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Workbook> Create()
    {
        var registry = new HandlerRegistry<Workbook>();
        registry.Register(new WriteExcelCellHandler());
        registry.Register(new EditExcelCellHandler());
        registry.Register(new GetExcelCellHandler());
        registry.Register(new ClearExcelCellHandler());
        return registry;
    }
}
