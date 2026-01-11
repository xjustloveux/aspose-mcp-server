using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.Properties;

/// <summary>
///     Registry for Excel properties operation handlers.
///     Provides a pre-configured registry with all properties handlers registered.
/// </summary>
public static class ExcelPropertiesHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Excel properties handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Workbook> Create()
    {
        var registry = new HandlerRegistry<Workbook>();
        registry.Register(new GetWorkbookPropertiesHandler());
        registry.Register(new SetWorkbookPropertiesHandler());
        registry.Register(new GetSheetPropertiesHandler());
        registry.Register(new EditSheetPropertiesHandler());
        registry.Register(new GetSheetInfoHandler());
        return registry;
    }
}
