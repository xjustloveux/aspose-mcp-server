using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.RowColumn;

/// <summary>
///     Registry for Excel row/column operation handlers.
///     Provides a pre-configured registry with all row/column handlers registered.
/// </summary>
public static class ExcelRowColumnHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Excel row/column handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Workbook> Create()
    {
        var registry = new HandlerRegistry<Workbook>();
        registry.Register(new InsertRowHandler());
        registry.Register(new DeleteRowHandler());
        registry.Register(new InsertColumnHandler());
        registry.Register(new DeleteColumnHandler());
        registry.Register(new InsertCellsHandler());
        registry.Register(new DeleteCellsHandler());
        return registry;
    }
}
