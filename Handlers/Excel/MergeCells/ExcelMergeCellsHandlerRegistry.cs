using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.MergeCells;

/// <summary>
///     Registry for Excel merge cells operation handlers.
///     Provides a pre-configured registry with all merge cells handlers registered.
/// </summary>
public static class ExcelMergeCellsHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Excel merge cells handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Workbook> Create()
    {
        var registry = new HandlerRegistry<Workbook>();
        registry.Register(new MergeExcelCellsHandler());
        registry.Register(new UnmergeExcelCellsHandler());
        registry.Register(new GetExcelMergedCellsHandler());
        return registry;
    }
}
