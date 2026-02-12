using Aspose.Slides;
using AsposeMcpServer.Core.ShapeDetailProviders.Providers;

namespace AsposeMcpServer.Core.ShapeDetailProviders;

/// <summary>
///     Factory for getting the appropriate shape detail provider for a given shape
/// </summary>
public static class ShapeDetailProviderFactory
{
    /// <summary>
    ///     Registered shape detail providers for handling different shape types.
    /// </summary>
    private static readonly List<IShapeDetailProvider> Providers =
    [
        new AutoShapeDetailProvider(),
        new PictureFrameDetailProvider(),
        new TableDetailProvider(),
        new ChartDetailProvider(),
        new SmartArtDetailProvider(),
        new GroupShapeDetailProvider(),
        new AudioFrameDetailProvider(),
        new VideoFrameDetailProvider(),
        new ConnectorDetailProvider()
    ];

    /// <summary>
    ///     Gets the appropriate provider for the given shape
    /// </summary>
    /// <param name="shape">The shape to find a provider for</param>
    /// <returns>The provider that can handle this shape, or null if no provider found</returns>
    public static IShapeDetailProvider? GetProvider(IShape shape)
    {
        return Providers.FirstOrDefault(p => p.CanHandle(shape));
    }

    /// <summary>
    ///     Gets shape details using the appropriate provider.
    /// </summary>
    /// <param name="shape">The shape to extract details from.</param>
    /// <param name="presentation">The presentation containing the shape.</param>
    /// <returns>A <see cref="ShapeTypeInfo" /> containing type name and details.</returns>
    public static ShapeTypeInfo GetShapeDetails(IShape shape, IPresentation presentation)
    {
        var provider = GetProvider(shape);
        if (provider != null) return new ShapeTypeInfo(provider.TypeName, provider.GetDetails(shape, presentation));

        return new ShapeTypeInfo(shape.GetType().Name, null);
    }
}
