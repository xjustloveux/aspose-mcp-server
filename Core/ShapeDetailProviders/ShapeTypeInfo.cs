using AsposeMcpServer.Core.ShapeDetailProviders.Details;

namespace AsposeMcpServer.Core.ShapeDetailProviders;

/// <summary>
///     Represents the type information and details for a PowerPoint shape.
/// </summary>
/// <param name="TypeName">The shape type name.</param>
/// <param name="Details">The shape details, or null if no provider found.</param>
public record ShapeTypeInfo(string TypeName, ShapeDetails? Details);
