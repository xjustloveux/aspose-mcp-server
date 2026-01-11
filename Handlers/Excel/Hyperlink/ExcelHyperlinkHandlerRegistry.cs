using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.Hyperlink;

/// <summary>
///     Registry for Excel hyperlink operation handlers.
///     Provides a pre-configured registry with all hyperlink handlers registered.
/// </summary>
public static class ExcelHyperlinkHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Excel hyperlink handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Workbook> Create()
    {
        var registry = new HandlerRegistry<Workbook>();
        registry.Register(new AddExcelHyperlinkHandler());
        registry.Register(new EditExcelHyperlinkHandler());
        registry.Register(new DeleteExcelHyperlinkHandler());
        registry.Register(new GetExcelHyperlinksHandler());
        return registry;
    }
}
