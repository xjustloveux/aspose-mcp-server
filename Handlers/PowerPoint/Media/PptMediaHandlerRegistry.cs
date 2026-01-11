using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Media;

/// <summary>
///     Registry for PowerPoint media operation handlers.
///     Provides a pre-configured registry with all media handlers registered.
/// </summary>
public static class PptMediaHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PowerPoint media handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Presentation> Create()
    {
        var registry = new HandlerRegistry<Presentation>();
        registry.Register(new AddAudioHandler());
        registry.Register(new DeleteAudioHandler());
        registry.Register(new AddVideoHandler());
        registry.Register(new DeleteVideoHandler());
        registry.Register(new SetPlaybackHandler());
        return registry;
    }
}
