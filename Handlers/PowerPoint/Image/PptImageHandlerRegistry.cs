using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Image;

/// <summary>
///     Registry for PowerPoint image operation handlers.
///     Provides a pre-configured registry with all image handlers registered.
/// </summary>
public static class PptImageHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PowerPoint image handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Presentation> Create()
    {
        var registry = new HandlerRegistry<Presentation>();
        registry.Register(new AddPptImageHandler());
        registry.Register(new EditPptImageHandler());
        registry.Register(new DeletePptImageHandler());
        registry.Register(new GetPptImagesHandler());
        registry.Register(new ExportSlidesHandler());
        registry.Register(new ExtractPptImageHandler());
        return registry;
    }
}
