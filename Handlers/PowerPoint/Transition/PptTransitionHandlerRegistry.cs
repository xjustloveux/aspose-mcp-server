using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Transition;

/// <summary>
///     Registry for PowerPoint transition operation handlers.
///     Provides a pre-configured registry with all transition handlers registered.
/// </summary>
public static class PptTransitionHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PowerPoint transition handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Presentation> Create()
    {
        var registry = new HandlerRegistry<Presentation>();
        registry.Register(new SetPptTransitionHandler());
        registry.Register(new GetPptTransitionHandler());
        registry.Register(new DeletePptTransitionHandler());
        return registry;
    }
}
