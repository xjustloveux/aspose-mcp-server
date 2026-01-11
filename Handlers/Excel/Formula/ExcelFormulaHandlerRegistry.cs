using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.Formula;

/// <summary>
///     Registry for Excel formula operation handlers.
///     Provides a pre-configured registry with all formula handlers registered.
/// </summary>
public static class ExcelFormulaHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Excel formula handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Workbook> Create()
    {
        var registry = new HandlerRegistry<Workbook>();
        registry.Register(new AddFormulaHandler());
        registry.Register(new GetFormulasHandler());
        registry.Register(new GetFormulaResultHandler());
        registry.Register(new CalculateFormulasHandler());
        registry.Register(new SetArrayFormulaHandler());
        registry.Register(new GetArrayFormulaHandler());
        return registry;
    }
}
