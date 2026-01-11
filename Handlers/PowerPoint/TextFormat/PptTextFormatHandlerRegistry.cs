using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.TextFormat;

/// <summary>
///     Registry for PowerPoint text format operation handlers.
///     Provides a pre-configured registry with all text format handlers registered.
/// </summary>
public static class PptTextFormatHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PowerPoint text format handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Presentation> Create()
    {
        var registry = new HandlerRegistry<Presentation>();
        registry.Register(new FormatPptTextHandler());
        return registry;
    }
}
