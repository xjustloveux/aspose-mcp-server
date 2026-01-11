using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.DataOperations;

/// <summary>
///     Registry for Excel data operations handlers.
///     Provides a pre-configured registry with all data operations handlers registered.
/// </summary>
public static class ExcelDataOperationsHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Excel data operations handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Workbook> Create()
    {
        var registry = new HandlerRegistry<Workbook>();
        registry.Register(new SortDataHandler());
        registry.Register(new FindReplaceHandler());
        registry.Register(new BatchWriteHandler());
        registry.Register(new GetContentHandler());
        registry.Register(new GetStatisticsHandler());
        registry.Register(new GetUsedRangeHandler());
        return registry;
    }
}
