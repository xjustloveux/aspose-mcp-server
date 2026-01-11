using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.Sheet;

/// <summary>
///     Registry for Excel sheet operation handlers.
///     Provides a pre-configured registry with all sheet handlers registered.
/// </summary>
public static class ExcelSheetHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Excel sheet handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Workbook> Create()
    {
        var registry = new HandlerRegistry<Workbook>();
        registry.Register(new AddExcelSheetHandler());
        registry.Register(new DeleteExcelSheetHandler());
        registry.Register(new GetExcelSheetsHandler());
        registry.Register(new RenameExcelSheetHandler());
        registry.Register(new MoveExcelSheetHandler());
        registry.Register(new CopyExcelSheetHandler());
        registry.Register(new HideExcelSheetHandler());
        return registry;
    }
}
