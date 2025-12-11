using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.SmartArt;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing PowerPoint SmartArt (add, manage nodes)
/// Merges: PptAddSmartArtTool, PptManageSmartArtNodesTool
/// </summary>
public class PptSmartArtTool : IAsposeTool
{
    public string Description => "Manage PowerPoint SmartArt: add or manage nodes";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = "Operation to perform: 'add', 'manage_nodes'",
                @enum = new[] { "add", "manage_nodes" }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path"
            },
            slideIndex = new
            {
                type = "number",
                description = "Slide index (0-based)"
            },
            shapeIndex = new
            {
                type = "number",
                description = "SmartArt shape index (0-based, required for manage_nodes)"
            },
            layout = new
            {
                type = "string",
                description = "Layout: BasicProcess, ContinuousCycle, Hierarchy, etc. (required for add)"
            },
            action = new
            {
                type = "string",
                description = "add|remove|rename|move (required for manage_nodes)"
            },
            targetPath = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Node path indices from root (e.g., [0,1], required for manage_nodes)"
            },
            text = new
            {
                type = "string",
                description = "Text for add/rename (required for manage_nodes)"
            },
            moveParentPath = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Destination parent path for move (optional, for manage_nodes)"
            },
            moveIndex = new
            {
                type = "number",
                description = "Insert index under new parent (optional, for manage_nodes)"
            },
            x = new
            {
                type = "number",
                description = "X position (optional, for add, default: 50)"
            },
            y = new
            {
                type = "number",
                description = "Y position (optional, for add, default: 50)"
            },
            width = new
            {
                type = "number",
                description = "Width (optional, for add, default: 400)"
            },
            height = new
            {
                type = "number",
                description = "Height (optional, for add, default: 300)"
            }
        },
        required = new[] { "operation", "path", "slideIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");

        return operation.ToLower() switch
        {
            "add" => await AddSmartArtAsync(arguments, path, slideIndex),
            "manage_nodes" => await ManageNodesAsync(arguments, path, slideIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> AddSmartArtAsync(JsonObject? arguments, string path, int slideIndex)
    {
        var layoutStr = arguments?["layout"]?.GetValue<string>() ?? throw new ArgumentException("layout is required for add operation");
        var x = arguments?["x"]?.GetValue<float?>() ?? 50;
        var y = arguments?["y"]?.GetValue<float?>() ?? 50;
        var width = arguments?["width"]?.GetValue<float?>() ?? 400;
        var height = arguments?["height"]?.GetValue<float?>() ?? 300;

        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        
        var layout = layoutStr.ToLower() switch
        {
            "basicprocess" => SmartArtLayoutType.BasicProcess,
            "continuouscycle" => SmartArtLayoutType.ContinuousCycle,
            "hierarchy" => SmartArtLayoutType.Hierarchy,
            "basicblocklist" => SmartArtLayoutType.BasicBlockList,
            "basicpyramid" => SmartArtLayoutType.BasicPyramid,
            "stackedlist" => SmartArtLayoutType.StackedList,
            "horizontalmultilevelhierarchy" => SmartArtLayoutType.HorizontalMultiLevelHierarchy,
            _ => SmartArtLayoutType.BasicProcess
        };
        slide.Shapes.AddSmartArt(x, y, width, height, layout);

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已新增 SmartArt ({layout}) 至投影片 {slideIndex}");
    }

    private async Task<string> ManageNodesAsync(JsonObject? arguments, string path, int slideIndex)
    {
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required for manage_nodes operation");
        var action = arguments?["action"]?.GetValue<string>()?.ToLower() ?? throw new ArgumentException("action is required for manage_nodes operation");
        var targetPath = arguments?["targetPath"]?.AsArray()?.Select(x => x?.GetValue<int>() ?? -1).ToArray();
        var text = arguments?["text"]?.GetValue<string>();
        var moveParentPath = arguments?["moveParentPath"]?.AsArray()?.Select(x => x?.GetValue<int>() ?? -1).ToArray();
        var moveIndex = arguments?["moveIndex"]?.GetValue<int?>();

        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex);

        if (shape is not ISmartArt smartArt)
        {
            throw new ArgumentException("指定的 shape 不是 SmartArt");
        }

        SmartArtNode? GetNode(int[]? pathArr)
        {
            if (pathArr == null || pathArr.Length == 0) return null;
            var node = smartArt.AllNodes[pathArr[0]] as SmartArtNode;
            for (int i = 1; node != null && i < pathArr.Length; i++)
            {
                if (pathArr[i] < 0 || pathArr[i] >= node.ChildNodes.Count) return null;
                node = node.ChildNodes[pathArr[i]] as SmartArtNode;
            }
            return node;
        }

        switch (action)
        {
            case "add":
                {
                    var parent = GetNode(targetPath) ?? smartArt.AllNodes.AddNode();
                    var newNode = parent.ChildNodes.AddNode();
                    newNode.TextFrame.Text = string.IsNullOrEmpty(text) ? "New Node" : text;
                    break;
                }
            case "remove":
                {
                    var node = GetNode(targetPath);
                    if (node == null) throw new ArgumentException("targetPath 無效");
                    node.Remove();
                    break;
                }
            case "rename":
                {
                    var node = GetNode(targetPath);
                    if (node == null) throw new ArgumentException("targetPath 無效");
                    node.TextFrame.Text = text ?? string.Empty;
                    break;
                }
            case "move":
                {
                    var node = GetNode(targetPath);
                    if (node == null) throw new ArgumentException("targetPath 無效");
                    var parent = GetNode(moveParentPath) ?? smartArt.AllNodes.AddNode();
                    var clone = parent.ChildNodes.AddNode();
                    clone.TextFrame.Text = node.TextFrame.Text;
                    foreach (SmartArtNode child in node.ChildNodes)
                    {
                        var childClone = clone.ChildNodes.AddNode();
                        childClone.TextFrame.Text = child.TextFrame.Text;
                    }
                    node.Remove();
                    break;
                }
            default:
                throw new ArgumentException("action 必須為 add/remove/rename/move");
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"SmartArt {action} 完成：slide {slideIndex}, shape {shapeIndex}");
    }
}

