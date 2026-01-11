using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Properties;

/// <summary>
///     Registry for PowerPoint properties operation handlers.
///     Provides a pre-configured registry with all properties handlers registered.
/// </summary>
public static class PptPropertiesHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PowerPoint properties handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Presentation> Create()
    {
        var registry = new HandlerRegistry<Presentation>();
        registry.Register(new GetPptPropertiesHandler());
        registry.Register(new SetPptPropertiesHandler());
        return registry;
    }
}
