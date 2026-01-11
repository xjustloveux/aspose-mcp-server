using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.PivotTable;

/// <summary>
///     Registry for Excel pivot table operation handlers.
///     Provides a pre-configured registry with all pivot table handlers registered.
/// </summary>
public static class ExcelPivotTableHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Excel pivot table handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Workbook> Create()
    {
        var registry = new HandlerRegistry<Workbook>();
        registry.Register(new AddExcelPivotTableHandler());
        registry.Register(new EditExcelPivotTableHandler());
        registry.Register(new DeleteExcelPivotTableHandler());
        registry.Register(new GetExcelPivotTablesHandler());
        registry.Register(new AddFieldExcelPivotTableHandler());
        registry.Register(new DeleteFieldExcelPivotTableHandler());
        registry.Register(new RefreshExcelPivotTableHandler());
        return registry;
    }
}
