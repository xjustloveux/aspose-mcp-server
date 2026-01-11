using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.FileOperations;

/// <summary>
///     Registry for PowerPoint file operation handlers.
///     Provides a pre-configured registry with all file operation handlers registered.
/// </summary>
public static class PptFileOperationsHandlerRegistry
{
    /// <summary>
    ///     Creates a new registry with all PowerPoint file operation handlers registered.
    /// </summary>
    /// <returns>A configured handler registry.</returns>
    public static HandlerRegistry<Presentation> Create()
    {
        var registry = new HandlerRegistry<Presentation>();
        registry.Register(new CreatePresentationHandler());
        registry.Register(new ConvertPresentationHandler());
        registry.Register(new MergePresentationsHandler());
        registry.Register(new SplitPresentationHandler());
        return registry;
    }
}
