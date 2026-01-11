using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.Chart;

/// <summary>
///     Registry for Excel chart operation handlers.
///     Provides a pre-configured registry with all chart handlers registered.
/// </summary>
public static class ExcelChartHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Excel chart handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Workbook> Create()
    {
        var registry = new HandlerRegistry<Workbook>();
        registry.Register(new AddExcelChartHandler());
        registry.Register(new EditExcelChartHandler());
        registry.Register(new DeleteExcelChartHandler());
        registry.Register(new GetExcelChartsHandler());
        registry.Register(new UpdateExcelChartDataHandler());
        registry.Register(new SetExcelChartPropertiesHandler());
        return registry;
    }
}
