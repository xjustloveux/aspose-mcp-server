using System.Text;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using Aspose.Slides.SmartArt;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint SmartArt (add, manage nodes)
///     Supports: add, manage_nodes
/// </summary>
public class PptSmartArtTool : IAsposeTool
{
    public string Description =>
        @"Manage PowerPoint SmartArt. Supports 2 operations: add, manage_nodes.

Usage examples:
- Add SmartArt: ppt_smart_art(operation='add', path='presentation.pptx', slideIndex=0, layout='BasicProcess', x=100, y=100, width=400, height=300, outputPath='output.pptx')
- Manage nodes: ppt_smart_art(operation='manage_nodes', path='presentation.pptx', slideIndex=0, shapeIndex=0, action='add', targetPath=[0], text='New Node', outputPath='output.pptx')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add a new SmartArt shape (required params: path, slideIndex, layout)
- 'manage_nodes': Manage SmartArt nodes (add, edit, delete) (required params: path, slideIndex, shapeIndex, action)",
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
                description = "Slide index (0-based, required for all operations)"
            },
            shapeIndex = new
            {
                type = "number",
                description = "Shape index (0-based, required for manage_nodes, refers to SmartArt shape index)"
            },
            layout = new
            {
                type = "string",
                description =
                    "SmartArt layout type (required for add operation, e.g., 'BasicProcess', 'Cycle', 'Hierarchy')"
            },
            x = new
            {
                type = "number",
                description = "X position (optional, for add operation, defaults to 100)"
            },
            y = new
            {
                type = "number",
                description = "Y position (optional, for add operation, defaults to 100)"
            },
            width = new
            {
                type = "number",
                description = "Width (optional, for add operation, defaults to 400)"
            },
            height = new
            {
                type = "number",
                description = "Height (optional, for add operation, defaults to 300)"
            },
            action = new
            {
                type = "string",
                description = "Node action: 'add', 'edit', 'delete' (required for manage_nodes operation)",
                @enum = new[] { "add", "edit", "delete" }
            },
            targetPath = new
            {
                type = "array",
                description =
                    "Array of indices to target node (required for manage_nodes, e.g., [0] for first node, [0,1] for second child of first node)"
            },
            text = new
            {
                type = "string",
                description = "Node text content (required for add/edit operations in manage_nodes)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to input path)"
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
    ///     Adds a SmartArt shape to a slide
    /// </summary>
    /// <param name="arguments">JSON arguments containing layout, x, y, width, height, outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <returns>Success message</returns>
    private Task<string> AddSmartArtAsync(JsonObject? arguments, string path, int slideIndex)
    {
        return Task.Run(() =>
        {
            var layoutStr = ArgumentHelper.GetString(arguments, "layout");
            var x = ArgumentHelper.GetFloat(arguments, "x", 100);
            var y = ArgumentHelper.GetFloat(arguments, "y", 100);
            var width = ArgumentHelper.GetFloat(arguments, "width", 400);
            var height = ArgumentHelper.GetFloat(arguments, "height", 300);

            // Parse layout string to SmartArtLayoutType enum
            if (!Enum.TryParse<SmartArtLayoutType>(layoutStr, true, out var layoutType))
                throw new ArgumentException(
                    $"Invalid SmartArt layout: '{layoutStr}'. Valid layouts include: BasicProcess, Cycle, Hierarchy, etc.");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

            slide.Shapes.AddSmartArt(x, y, width, height, layoutType);

            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            presentation.Save(outputPath, SaveFormat.Pptx);

            return $"SmartArt added successfully to slide {slideIndex} with layout '{layoutStr}' - {outputPath}";
        });
    }

    /// <summary>
    ///     Manages SmartArt nodes (add, edit, delete)
    /// </summary>
    /// <param name="arguments">JSON arguments containing shapeIndex, action, targetPath, text, outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <returns>Success message</returns>
    private Task<string> ManageNodesAsync(JsonObject? arguments, string path, int slideIndex)
    {
        return Task.Run(() =>
        {
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
            var action = ArgumentHelper.GetString(arguments, "action");
            var targetPathArray = ArgumentHelper.GetArray(arguments, "targetPath");
            var text = ArgumentHelper.GetStringNullable(arguments, "text");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

            PowerPointHelper.ValidateShapeIndex(shapeIndex, slide);
            var shape = slide.Shapes[shapeIndex];

            if (shape is not ISmartArt smartArt)
                throw new ArgumentException(
                    $"Shape at index {shapeIndex} is not a SmartArt shape. It is a {shape.GetType().Name}.");

            // Convert targetPath array to list of indices
            var targetPath = new List<int>();
            foreach (var item in targetPathArray)
                if (item != null && int.TryParse(item.ToString(), out var index))
                    targetPath.Add(index);
                else
                    throw new ArgumentException(
                        $"Invalid targetPath: all elements must be integers. Found: {item}");

            if (targetPath.Count == 0) throw new ArgumentException("targetPath must contain at least one index");

            // Navigate to target node
            ISmartArtNode? targetNode = null;
            var currentNode = smartArt.AllNodes[targetPath[0]];

            if (targetPath.Count == 1)
            {
                targetNode = currentNode;
            }
            else
            {
                for (var i = 1; i < targetPath.Count; i++)
                {
                    var childIndex = targetPath[i];
                    if (childIndex < 0 || childIndex >= currentNode.ChildNodes.Count)
                        throw new ArgumentException(
                            $"Child index {childIndex} is out of range (node has {currentNode.ChildNodes.Count} children)");

                    currentNode = currentNode.ChildNodes[childIndex];
                }

                targetNode = currentNode;
            }

            if (targetNode == null) throw new InvalidOperationException("Unable to locate target node");

            var sb = new StringBuilder();

            switch (action.ToLower())
            {
                case "add":
                    if (string.IsNullOrEmpty(text))
                        throw new ArgumentException("text parameter is required for 'add' action");

                    var newNode = targetNode.ChildNodes.AddNode();
                    newNode.TextFrame.Text = text;
                    sb.AppendLine("Node added successfully to SmartArt");
                    sb.AppendLine($"Target path: [{string.Join(",", targetPath)}]");
                    sb.AppendLine($"New node text: {text}");
                    break;

                case "edit":
                    if (string.IsNullOrEmpty(text))
                        throw new ArgumentException("text parameter is required for 'edit' action");

                    targetNode.TextFrame.Text = text;
                    sb.AppendLine("Node edited successfully");
                    sb.AppendLine($"Target path: [{string.Join(",", targetPath)}]");
                    sb.AppendLine($"New text: {text}");
                    break;

                case "delete":
                    // Check if it's a root node (root nodes are directly in AllNodes, not in ChildNodes)
                    if (targetPath.Count == 1) throw new InvalidOperationException("Cannot delete root node");

                    var nodeText = targetNode.TextFrame?.Text ?? "(empty)";
                    targetNode.Remove();
                    sb.AppendLine("Node deleted successfully");
                    sb.AppendLine($"Deleted node path: [{string.Join(",", targetPath)}]");
                    sb.AppendLine($"Deleted node text: {nodeText}");
                    break;

                default:
                    throw new ArgumentException($"Unknown action: {action}. Valid actions: add, edit, delete");
            }

            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            presentation.Save(outputPath, SaveFormat.Pptx);

            sb.AppendLine($"Output: {outputPath}");
            return sb.ToString();
        });
    }
}