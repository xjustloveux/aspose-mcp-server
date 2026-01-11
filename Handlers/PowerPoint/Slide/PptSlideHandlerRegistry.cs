using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Slide;

/// <summary>
///     Registry for PowerPoint slide operation handlers.
///     Provides a pre-configured registry with all slide handlers registered.
/// </summary>
public static class PptSlideHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PowerPoint slide handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Presentation> Create()
    {
        var registry = new HandlerRegistry<Presentation>();
        registry.Register(new AddPptSlideHandler());
        registry.Register(new DeletePptSlideHandler());
        registry.Register(new GetPptSlidesInfoHandler());
        registry.Register(new MovePptSlideHandler());
        registry.Register(new DuplicatePptSlideHandler());
        registry.Register(new HidePptSlidesHandler());
        registry.Register(new ClearPptSlideHandler());
        registry.Register(new EditPptSlideHandler());
        return registry;
    }
}
