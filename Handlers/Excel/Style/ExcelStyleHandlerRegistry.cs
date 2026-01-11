using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.Style;

/// <summary>
///     Registry for Excel style operation handlers.
///     Provides a pre-configured registry with all style handlers registered.
/// </summary>
public static class ExcelStyleHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Excel style handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Workbook> Create()
    {
        var registry = new HandlerRegistry<Workbook>();
        registry.Register(new FormatCellsHandler());
        registry.Register(new GetCellFormatHandler());
        registry.Register(new CopySheetFormatHandler());
        return registry;
    }
}
