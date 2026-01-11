using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.DataValidation;

/// <summary>
///     Registry for Excel data validation operation handlers.
///     Provides a pre-configured registry with all data validation handlers registered.
/// </summary>
public static class ExcelDataValidationHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Excel data validation handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Workbook> Create()
    {
        var registry = new HandlerRegistry<Workbook>();
        registry.Register(new AddExcelDataValidationHandler());
        registry.Register(new EditExcelDataValidationHandler());
        registry.Register(new DeleteExcelDataValidationHandler());
        registry.Register(new GetExcelDataValidationsHandler());
        registry.Register(new SetMessagesExcelDataValidationHandler());
        return registry;
    }
}
