using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using Aspose.Slides.SmartArt;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint SmartArt (add, manage nodes)
///     Merges: PptAddSmartArtTool, PptManageSmartArtNodesTool
/// </summary>
public class PptSmartArtTool : IAsposeTool
{
    public string Description => @"Manage PowerPoint SmartArt. Supports 2 operations: add, manage_nodes.

Usage examples:
- Add SmartArt: ppt_smart_art(operation='add', path='presentation.pptx', slideIndex=0, layout='BasicProcess', x=100, y=100, width=400, height=300)
- Manage nodes: ppt_smart_art(operation='manage_nodes', path='presentation.pptx', slideIndex=0, shapeIndex=0, action='add', targetPath=[0], text='New Node')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add SmartArt shape (required params: path, slideIndex, layout, x, y, width, height)
- 'manage_nodes': Manage SmartArt nodes (required params: path, slideIndex, shapeIndex, action, targetPath)",
                @enum = new[] { "add", "manage_nodes" }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path (required for all operations)"
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
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for add/edit/delete operations, defaults to input path)"
            }
        },
        required = new[] { "operation", "path", "slideIndex" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");

        return operation.ToLower() switch
        {
            "add" => await AddSmartArtAsync(arguments, path, slideIndex),
            "manage_nodes" => await ManageNodesAsync(arguments, path, slideIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a SmartArt diagram to a slide
    /// </summary>
    /// <param name="arguments">JSON arguments containing smartArtType, optional x, y, width, height, outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> AddSmartArtAsync(JsonObject? arguments, string path, int slideIndex)
    {
        var layoutStr = ArgumentHelper.GetString(arguments, "layout");
        var x = ArgumentHelper.GetFloat(arguments, "x", 50);
        var y = ArgumentHelper.GetFloat(arguments, "y", 50);
        var width = ArgumentHelper.GetFloat(arguments, "width", 400);
        var height = ArgumentHelper.GetFloat(arguments, "height", 300);

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

        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        presentation.Save(outputPath, SaveFormat.Pptx);
        return await Task.FromResult($"SmartArt ({layout}) added to slide {slideIndex}");
    }

    /// <summary>
    ///     Manages SmartArt nodes (add, edit, delete)
    /// </summary>
    /// <param name="arguments">JSON arguments containing smartArtIndex, action, optional nodeIndex, text, outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <returns>Success message</returns>
    private async Task<string> ManageNodesAsync(JsonObject? arguments, string path, int slideIndex)
    {
        var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
        var action = ArgumentHelper.GetString(arguments, "action").ToLower();
        var targetPathArray = ArgumentHelper.GetArray(arguments, "targetPath", false);
        var targetPath = targetPathArray?.Select(x => x?.GetValue<int>() ?? -1).ToArray();
        var text = ArgumentHelper.GetStringNullable(arguments, "text");
        var moveParentPathArray = ArgumentHelper.GetArray(arguments, "moveParentPath", false);
        var moveParentPath = moveParentPathArray?.Select(x => x?.GetValue<int>() ?? -1).ToArray();
        _ = ArgumentHelper.GetIntNullable(arguments, "moveIndex");

        using var presentation = new Presentation(path);
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex);

        if (shape is not ISmartArt smartArt) throw new ArgumentException("The specified shape is not a SmartArt");

        SmartArtNode? GetNode(int[]? pathArr)
        {
            if (pathArr == null || pathArr.Length == 0) return null;
            var node = smartArt.AllNodes[pathArr[0]] as SmartArtNode;
            for (var i = 1; node != null && i < pathArr.Length; i++)
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
                if (node == null) throw new ArgumentException("targetPath is invalid");
                node.Remove();
                break;
            }
            case "rename":
            {
                var node = GetNode(targetPath);
                if (node == null) throw new ArgumentException("targetPath is invalid");
                node.TextFrame.Text = text ?? string.Empty;
                break;
            }
            case "move":
            {
                var node = GetNode(targetPath);
                if (node == null) throw new ArgumentException("targetPath is invalid");
                var parent = GetNode(moveParentPath) ?? smartArt.AllNodes.AddNode();
                var clone = parent.ChildNodes.AddNode();
                clone.TextFrame.Text = node.TextFrame.Text;
                foreach (var child in node.ChildNodes.Cast<SmartArtNode>())
                {
                    var childClone = clone.ChildNodes.AddNode();
                    childClone.TextFrame.Text = child.TextFrame.Text;
                }

                node.Remove();
                break;
            }
            default:
                throw new ArgumentException("action must be one of: add/remove/rename/move");
        }

        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        presentation.Save(outputPath, SaveFormat.Pptx);
        return await Task.FromResult($"SmartArt {action} completed: slide {slideIndex}, shape {shapeIndex}");
    }
}