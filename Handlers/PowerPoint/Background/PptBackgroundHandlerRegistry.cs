using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Background;

/// <summary>
///     Registry for PowerPoint background operation handlers.
///     Provides a pre-configured registry with all background handlers registered.
/// </summary>
public static class PptBackgroundHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PowerPoint background handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Presentation> Create()
    {
        var registry = new HandlerRegistry<Presentation>();
        registry.Register(new SetPptBackgroundHandler());
        registry.Register(new GetPptBackgroundHandler());
        return registry;
    }
}
