using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.FileOperations;

/// <summary>
///     Registry for Excel file operation handlers.
///     Provides a pre-configured registry with all file operation handlers registered.
/// </summary>
public static class ExcelFileOperationsHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Excel file operation handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Workbook> Create()
    {
        var registry = new HandlerRegistry<Workbook>();
        registry.Register(new CreateWorkbookHandler());
        registry.Register(new ConvertWorkbookHandler());
        registry.Register(new MergeWorkbooksHandler());
        registry.Register(new SplitWorkbookHandler());
        return registry;
    }
}
