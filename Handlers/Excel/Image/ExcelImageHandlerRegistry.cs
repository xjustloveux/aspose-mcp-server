using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.Image;

/// <summary>
///     Registry for Excel image operation handlers.
///     Provides a pre-configured registry with all image handlers registered.
/// </summary>
public static class ExcelImageHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Excel image handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Workbook> Create()
    {
        var registry = new HandlerRegistry<Workbook>();
        registry.Register(new AddExcelImageHandler());
        registry.Register(new DeleteExcelImageHandler());
        registry.Register(new GetExcelImagesHandler());
        registry.Register(new ExtractExcelImageHandler());
        return registry;
    }
}
