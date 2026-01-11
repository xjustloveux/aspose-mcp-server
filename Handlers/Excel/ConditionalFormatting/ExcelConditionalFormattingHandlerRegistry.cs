using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.ConditionalFormatting;

/// <summary>
///     Registry for Excel conditional formatting operation handlers.
///     Provides a pre-configured registry with all conditional formatting handlers registered.
/// </summary>
public static class ExcelConditionalFormattingHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Excel conditional formatting handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Workbook> Create()
    {
        var registry = new HandlerRegistry<Workbook>();
        registry.Register(new AddExcelConditionalFormattingHandler());
        registry.Register(new EditExcelConditionalFormattingHandler());
        registry.Register(new DeleteExcelConditionalFormattingHandler());
        registry.Register(new GetExcelConditionalFormattingsHandler());
        return registry;
    }
}
