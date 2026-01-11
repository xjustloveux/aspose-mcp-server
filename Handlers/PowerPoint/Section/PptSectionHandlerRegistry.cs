using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Section;

/// <summary>
///     Registry for PowerPoint section operation handlers.
///     Provides a pre-configured registry with all section handlers registered.
/// </summary>
public static class PptSectionHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PowerPoint section handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Presentation> Create()
    {
        var registry = new HandlerRegistry<Presentation>();
        registry.Register(new AddPptSectionHandler());
        registry.Register(new RenamePptSectionHandler());
        registry.Register(new DeletePptSectionHandler());
        registry.Register(new GetPptSectionsHandler());
        return registry;
    }
}
