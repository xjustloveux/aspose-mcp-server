using Aspose.Slides;
using AsposeMcpServer.Core.ShapeDetailProviders.Details;

namespace AsposeMcpServer.Core.ShapeDetailProviders.Providers;

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
    public ShapeDetails? GetDetails(IShape shape, IPresentation presentation)
    {
        if (shape is not IGroupShape groupShape)
            return null;

        List<GroupChildShapeInfo> childShapeList = [];
        for (var i = 0; i < groupShape.Shapes.Count; i++)
        {
            var s = groupShape.Shapes[i];
            var provider = ShapeDetailProviderFactory.GetProvider(s);
            childShapeList.Add(new GroupChildShapeInfo
            {
                Index = i,
                Name = string.IsNullOrEmpty(s.Name) ? null : s.Name,
                Type = provider?.TypeName ?? s.GetType().Name,
                Position = new ShapePositionInfo { X = s.X, Y = s.Y },
                Size = new ShapeSizeInfo { Width = s.Width, Height = s.Height }
            });
        }

        return new GroupShapeDetails
        {
            ChildCount = groupShape.Shapes.Count,
            Children = childShapeList.Count > 0 ? childShapeList : null
        };
    }
}
