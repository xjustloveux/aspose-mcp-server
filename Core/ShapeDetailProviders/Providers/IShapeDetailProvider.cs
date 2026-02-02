using Aspose.Slides;
using AsposeMcpServer.Core.ShapeDetailProviders.Details;

namespace AsposeMcpServer.Core.ShapeDetailProviders.Providers;

/// <summary>
///     Interface for extracting type-specific properties from PowerPoint shapes
/// </summary>
public interface IShapeDetailProvider
{
    /// <summary>
    ///     Gets the friendly type name for this shape type
    /// </summary>
    string TypeName { get; }

    /// <summary>
    ///     Checks if this provider can handle the given shape
    /// </summary>
    /// <param name="shape">The shape to check</param>
    /// <returns>True if this provider can extract details from the shape</returns>
    bool CanHandle(IShape shape);

    /// <summary>
    ///     Extracts type-specific properties from the shape
    /// </summary>
    /// <param name="shape">The shape to extract properties from</param>
    /// <param name="presentation">The presentation containing the shape (for context like slide references)</param>
    /// <returns>A strongly-typed detail record, or null if no properties to extract</returns>
    ShapeDetails? GetDetails(IShape shape, IPresentation presentation);
}
