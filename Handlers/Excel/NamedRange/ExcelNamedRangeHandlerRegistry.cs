using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.NamedRange;

/// <summary>
///     Registry for Excel named range operation handlers.
///     Provides a pre-configured registry with all named range handlers registered.
/// </summary>
public static class ExcelNamedRangeHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Excel named range handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Workbook> Create()
    {
        var registry = new HandlerRegistry<Workbook>();
        registry.Register(new AddExcelNamedRangeHandler());
        registry.Register(new DeleteExcelNamedRangeHandler());
        registry.Register(new GetExcelNamedRangesHandler());
        return registry;
    }
}
