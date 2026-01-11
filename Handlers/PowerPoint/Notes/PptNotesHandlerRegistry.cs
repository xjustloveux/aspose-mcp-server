using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Notes;

/// <summary>
///     Registry for PowerPoint notes operation handlers.
///     Provides a pre-configured registry with all notes handlers registered.
/// </summary>
public static class PptNotesHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PowerPoint notes handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Presentation> Create()
    {
        var registry = new HandlerRegistry<Presentation>();
        registry.Register(new SetNotesHandler());
        registry.Register(new GetNotesHandler());
        registry.Register(new ClearNotesHandler());
        registry.Register(new SetNotesHeaderFooterHandler());
        return registry;
    }
}
