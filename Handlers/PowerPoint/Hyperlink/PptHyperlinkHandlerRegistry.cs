using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Hyperlink;

/// <summary>
///     Registry for PowerPoint hyperlink operation handlers.
///     Provides a pre-configured registry with all hyperlink handlers registered.
/// </summary>
public static class PptHyperlinkHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PowerPoint hyperlink handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Presentation> Create()
    {
        var registry = new HandlerRegistry<Presentation>();
        registry.Register(new AddPptHyperlinkHandler());
        registry.Register(new EditPptHyperlinkHandler());
        registry.Register(new DeletePptHyperlinkHandler());
        registry.Register(new GetPptHyperlinksHandler());
        return registry;
    }
}
