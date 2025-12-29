using Aspose.Slides;
using Aspose.Slides.SmartArt;

namespace AsposeMcpServer.Core.ShapeDetailProviders;

/// <summary>
///     Provider for extracting details from SmartArt elements
/// </summary>
public class SmartArtDetailProvider : IShapeDetailProvider
{
    public string TypeName => "SmartArt";

    public bool CanHandle(IShape shape)
    {
        return shape is ISmartArt;
    }

    public object? GetDetails(IShape shape, IPresentation presentation)
    {
        if (shape is not ISmartArt smartArt)
            return null;

        // Get all nodes text
        var nodeTexts = GetNodeTexts(smartArt.AllNodes);

        return new
        {
            layout = smartArt.Layout.ToString(),
            quickStyle = smartArt.QuickStyle.ToString(),
            colorStyle = smartArt.ColorStyle.ToString(),
            isReversed = smartArt.IsReversed,
            nodeCount = smartArt.AllNodes.Count,
            nodes = nodeTexts.Length > 0 ? nodeTexts : null
        };
    }

    private static object[] GetNodeTexts(ISmartArtNodeCollection nodes)
    {
        var result = new List<object>();
        foreach (var node in nodes)
        {
            var text = node.TextFrame?.Text;
            var childTexts = node.ChildNodes.Count > 0 ? GetNodeTexts(node.ChildNodes) : null;

            result.Add(new
            {
                text,
                level = node.Level,
                isHidden = node.IsHidden,
                childCount = node.ChildNodes.Count,
                children = childTexts
            });
        }

        return result.ToArray();
    }
}