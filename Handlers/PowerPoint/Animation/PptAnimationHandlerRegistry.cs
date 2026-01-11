using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Animation;

/// <summary>
///     Registry for PowerPoint animation operation handlers.
///     Provides a pre-configured registry with all animation handlers registered.
/// </summary>
public static class PptAnimationHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PowerPoint animation handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Presentation> Create()
    {
        var registry = new HandlerRegistry<Presentation>();
        registry.Register(new AddPptAnimationHandler());
        registry.Register(new EditPptAnimationHandler());
        registry.Register(new DeletePptAnimationHandler());
        registry.Register(new GetPptAnimationsHandler());
        return registry;
    }
}
