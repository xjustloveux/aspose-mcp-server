using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.SmartArt;

/// <summary>
///     Registry for PowerPoint SmartArt operation handlers.
///     Provides a pre-configured registry with all SmartArt handlers registered.
/// </summary>
public static class PptSmartArtHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PowerPoint SmartArt handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Presentation> Create()
    {
        var registry = new HandlerRegistry<Presentation>();
        registry.Register(new AddSmartArtHandler());
        registry.Register(new ManageSmartArtNodesHandler());
        return registry;
    }
}
