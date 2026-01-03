using Aspose.Slides;
using Aspose.Slides.SmartArt;

namespace AsposeMcpServer.Core.ShapeDetailProviders;

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
    public object? GetDetails(IShape shape, IPresentation presentation)
    {
        if (shape is not ISmartArt smartArt)
            return null;

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

    /// <summary>
    ///     Recursively extracts text and hierarchy information from SmartArt nodes.
    /// </summary>
    /// <param name="nodes">The SmartArt node collection to process.</param>
    /// <returns>An array of objects containing node text and hierarchy details.</returns>
    private static object[] GetNodeTexts(ISmartArtNodeCollection nodes)
    {
        List<object> result = [];
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