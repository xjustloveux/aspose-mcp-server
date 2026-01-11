using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.PrintSettings;

/// <summary>
///     Registry for Excel print settings operation handlers.
///     Provides a pre-configured registry with all print settings handlers registered.
/// </summary>
public static class ExcelPrintSettingsHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Excel print settings handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Workbook> Create()
    {
        var registry = new HandlerRegistry<Workbook>();
        registry.Register(new SetPrintAreaHandler());
        registry.Register(new SetPrintTitlesHandler());
        registry.Register(new SetPageSetupHandler());
        registry.Register(new SetAllPrintSettingsHandler());
        return registry;
    }
}
