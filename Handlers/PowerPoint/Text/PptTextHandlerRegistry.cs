using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Text;

/// <summary>
///     Registry for PowerPoint text operation handlers.
///     Provides a pre-configured registry with all text handlers registered.
/// </summary>
public static class PptTextHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PowerPoint text handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Presentation> Create()
    {
        var registry = new HandlerRegistry<Presentation>();
        registry.Register(new AddPptTextHandler());
        registry.Register(new EditPptTextHandler());
        registry.Register(new ReplacePptTextHandler());
        return registry;
    }
}
