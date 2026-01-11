using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.Filter;

/// <summary>
///     Registry for Excel filter operation handlers.
///     Provides a pre-configured registry with all filter handlers registered.
/// </summary>
public static class ExcelFilterHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Excel filter handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Workbook> Create()
    {
        var registry = new HandlerRegistry<Workbook>();
        registry.Register(new ApplyExcelFilterHandler());
        registry.Register(new RemoveExcelFilterHandler());
        registry.Register(new FilterByValueExcelHandler());
        registry.Register(new GetExcelFilterStatusHandler());
        return registry;
    }
}
