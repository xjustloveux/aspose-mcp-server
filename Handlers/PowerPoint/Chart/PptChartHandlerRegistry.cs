using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Chart;

/// <summary>
///     Registry for PowerPoint chart operation handlers.
///     Provides a pre-configured registry with all chart handlers registered.
/// </summary>
public static class PptChartHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PowerPoint chart handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Presentation> Create()
    {
        var registry = new HandlerRegistry<Presentation>();
        registry.Register(new AddPptChartHandler());
        registry.Register(new EditPptChartHandler());
        registry.Register(new DeletePptChartHandler());
        registry.Register(new GetPptChartDataHandler());
        registry.Register(new UpdatePptChartDataHandler());
        return registry;
    }
}
