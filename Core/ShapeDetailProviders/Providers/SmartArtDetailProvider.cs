using Aspose.Slides;
using Aspose.Slides.SmartArt;
using AsposeMcpServer.Core.ShapeDetailProviders.Details;

namespace AsposeMcpServer.Core.ShapeDetailProviders.Providers;

/// <summary>
///     Provider for extracting details from SmartArt elements
/// </summary>
public class SmartArtDetailProvider : IShapeDetailProvider
{
    /// <inheritdoc />
    public string TypeName => "SmartArt";

    /// <inheritdoc />
    public bool CanHandle(IShape shape)
    {
        return shape is ISmartArt;
    }

    /// <inheritdoc />
    public ShapeDetails? GetDetails(IShape shape, IPresentation presentation)
    {
        if (shape is not ISmartArt smartArt)
            return null;

        var nodeInfos = GetNodeInfos(smartArt.AllNodes);

        return new SmartArtDetails
        {
            Layout = smartArt.Layout.ToString(),
            QuickStyle = smartArt.QuickStyle.ToString(),
            ColorStyle = smartArt.ColorStyle.ToString(),
            IsReversed = smartArt.IsReversed,
            NodeCount = smartArt.AllNodes.Count,
            Nodes = nodeInfos.Count > 0 ? nodeInfos : null
        };
    }

    /// <summary>
    ///     Recursively extracts hierarchy information from SmartArt nodes.
    /// </summary>
    /// <param name="nodes">The SmartArt node collection to process.</param>
    /// <returns>A list of node information records.</returns>
    private static List<SmartArtNodeInfo> GetNodeInfos(ISmartArtNodeCollection nodes)
    {
        List<SmartArtNodeInfo> result = [];
        foreach (var node in nodes)
        {
            var children = node.ChildNodes.Count > 0 ? GetNodeInfos(node.ChildNodes) : null;

            result.Add(new SmartArtNodeInfo
            {
                Text = node.TextFrame?.Text,
                Level = node.Level,
                IsHidden = node.IsHidden,
                ChildCount = node.ChildNodes.Count,
                Children = children
            });
        }

        return result;
    }
}
