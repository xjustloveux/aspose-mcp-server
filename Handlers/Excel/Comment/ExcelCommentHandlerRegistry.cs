using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.Comment;

/// <summary>
///     Registry for Excel comment operation handlers.
///     Provides a pre-configured registry with all comment handlers registered.
/// </summary>
public static class ExcelCommentHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Excel comment handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Workbook> Create()
    {
        var registry = new HandlerRegistry<Workbook>();
        registry.Register(new AddExcelCommentHandler());
        registry.Register(new EditExcelCommentHandler());
        registry.Register(new DeleteExcelCommentHandler());
        registry.Register(new GetExcelCommentsHandler());
        return registry;
    }
}
