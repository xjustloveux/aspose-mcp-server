using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.PageSetup;

/// <summary>
///     Registry for PowerPoint page setup operation handlers.
/// </summary>
public static class PptPageSetupHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PowerPoint page setup handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Presentation> Create()
    {
        var registry = new HandlerRegistry<Presentation>();
        registry.Register(new SetSlideSizeHandler());
        registry.Register(new SetSlideOrientationHandler());
        registry.Register(new SetFooterHandler());
        registry.Register(new SetSlideNumberingHandler());
        return registry;
    }
}
