using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.Range;

/// <summary>
///     Registry for Excel range operation handlers.
///     Provides a pre-configured registry with all range handlers registered.
/// </summary>
public static class ExcelRangeHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Excel range handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Workbook> Create()
    {
        var registry = new HandlerRegistry<Workbook>();
        registry.Register(new WriteExcelRangeHandler());
        registry.Register(new EditExcelRangeHandler());
        registry.Register(new GetExcelRangeHandler());
        registry.Register(new ClearExcelRangeHandler());
        registry.Register(new CopyExcelRangeHandler());
        registry.Register(new MoveExcelRangeHandler());
        registry.Register(new CopyFormatExcelRangeHandler());
        return registry;
    }
}
