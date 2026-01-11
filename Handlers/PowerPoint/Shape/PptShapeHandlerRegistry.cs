using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Shape;

/// <summary>
///     Registry for PowerPoint shape operation handlers.
///     Provides a pre-configured registry with all shape handlers registered.
/// </summary>
public static class PptShapeHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PowerPoint shape handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Presentation> Create()
    {
        var registry = new HandlerRegistry<Presentation>();
        registry.Register(new GetPptShapesHandler());
        registry.Register(new GetPptShapeDetailsHandler());
        registry.Register(new DeletePptShapeHandler());
        registry.Register(new EditPptShapeHandler());
        registry.Register(new SetPptShapeFormatHandler());
        registry.Register(new ClearPptShapeFormatHandler());
        registry.Register(new GroupPptShapesHandler());
        registry.Register(new UngroupPptShapesHandler());
        registry.Register(new CopyPptShapeHandler());
        registry.Register(new ReorderPptShapeHandler());
        registry.Register(new AlignPptShapesHandler());
        registry.Register(new FlipPptShapeHandler());
        return registry;
    }
}
