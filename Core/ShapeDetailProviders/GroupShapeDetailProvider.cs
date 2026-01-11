using Aspose.Slides;

namespace AsposeMcpServer.Core.ShapeDetailProviders;

/// <summary>
///     Provider for extracting details from GroupShape elements
/// </summary>
public class GroupShapeDetailProvider : IShapeDetailProvider
{
    /// <inheritdoc />
    public string TypeName => "Group";

    /// <inheritdoc />
    public bool CanHandle(IShape shape)
    {
        return shape is IGroupShape;
    }

    /// <inheritdoc />
    public object? GetDetails(IShape shape, IPresentation presentation)
    {
        if (shape is not IGroupShape groupShape)
            return null;

        List<object> childShapeList = [];
        for (var i = 0; i < groupShape.Shapes.Count; i++)
        {
            var s = groupShape.Shapes[i];
            var provider = ShapeDetailProviderFactory.GetProvider(s);
            childShapeList.Add(new
            {
                index = i,
                type = provider?.TypeName ?? s.GetType().Name,
                position = new { x = s.X, y = s.Y },
                size = new { width = s.Width, height = s.Height }
            });
        }

        return new
        {
            childCount = groupShape.Shapes.Count,
            children = childShapeList.Count > 0 ? childShapeList.ToArray() : null
        };
    }
}
