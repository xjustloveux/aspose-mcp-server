using Aspose.Slides;

namespace AsposeMcpServer.Core.ShapeDetailProviders;

/// <summary>
///     Provider for extracting details from GroupShape elements
/// </summary>
public class GroupShapeDetailProvider : IShapeDetailProvider
{
    public string TypeName => "Group";

    public bool CanHandle(IShape shape)
    {
        return shape is IGroupShape;
    }

    public object? GetDetails(IShape shape, IPresentation presentation)
    {
        if (shape is not IGroupShape groupShape)
            return null;

        // Extract child shapes manually since IShapeCollection may not support LINQ Select with index
        var childShapeList = new List<object>();
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