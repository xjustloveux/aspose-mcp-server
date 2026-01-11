using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Layout;

/// <summary>
///     Registry for PowerPoint layout operation handlers.
///     Provides a pre-configured registry with all layout handlers registered.
/// </summary>
public static class PptLayoutHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PowerPoint layout handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Presentation> Create()
    {
        var registry = new HandlerRegistry<Presentation>();
        registry.Register(new SetLayoutHandler());
        registry.Register(new GetLayoutsHandler());
        registry.Register(new GetMastersHandler());
        registry.Register(new ApplyMasterHandler());
        registry.Register(new ApplyLayoutRangeHandler());
        registry.Register(new ApplyThemeHandler());
        return registry;
    }
}
