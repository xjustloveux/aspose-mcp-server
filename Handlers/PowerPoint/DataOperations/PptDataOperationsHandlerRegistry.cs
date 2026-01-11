using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.DataOperations;

/// <summary>
///     Registry for PowerPoint data operations handlers.
///     Provides a pre-configured registry with all data operations handlers registered.
/// </summary>
public static class PptDataOperationsHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PowerPoint data operations handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Presentation> Create()
    {
        var registry = new HandlerRegistry<Presentation>();
        registry.Register(new GetStatisticsHandler());
        registry.Register(new GetContentHandler());
        registry.Register(new GetSlideDetailsHandler());
        return registry;
    }
}
