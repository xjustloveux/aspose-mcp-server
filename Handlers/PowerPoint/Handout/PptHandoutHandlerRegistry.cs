using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Handout;

/// <summary>
///     Registry for PowerPoint handout operation handlers.
///     Provides a pre-configured registry with all handout handlers registered.
/// </summary>
public static class PptHandoutHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PowerPoint handout handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Presentation> Create()
    {
        var registry = new HandlerRegistry<Presentation>();
        registry.Register(new SetHeaderFooterPptHandoutHandler());
        return registry;
    }
}
