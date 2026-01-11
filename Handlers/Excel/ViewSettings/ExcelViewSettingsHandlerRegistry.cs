using Aspose.Cells;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Excel.ViewSettings;

/// <summary>
///     Registry for Excel view settings operation handlers.
/// </summary>
public static class ExcelViewSettingsHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all Excel view settings handlers registered.
    /// </summary>
    public static HandlerRegistry<Workbook> Create()
    {
        var registry = new HandlerRegistry<Workbook>();
        registry.Register(new SetZoomExcelViewHandler());
        registry.Register(new SetGridlinesExcelViewHandler());
        registry.Register(new SetHeadersExcelViewHandler());
        registry.Register(new SetZeroValuesExcelViewHandler());
        registry.Register(new SetColumnWidthExcelViewHandler());
        registry.Register(new SetRowHeightExcelViewHandler());
        registry.Register(new SetBackgroundExcelViewHandler());
        registry.Register(new SetTabColorExcelViewHandler());
        registry.Register(new SetAllExcelViewHandler());
        registry.Register(new FreezePanesExcelViewHandler());
        registry.Register(new SplitWindowExcelViewHandler());
        registry.Register(new AutoFitColumnExcelViewHandler());
        registry.Register(new AutoFitRowExcelViewHandler());
        registry.Register(new ShowFormulasExcelViewHandler());
        return registry;
    }
}
