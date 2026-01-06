using System.ComponentModel;
using System.Text;
using System.Text.Json;
using Aspose.Slides;
using Aspose.Slides.SmartArt;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint SmartArt (add, manage nodes)
///     Supports: add, manage_nodes
/// </summary>
[McpServerToolType]
public class PptSmartArtTool
{
    /// <summary>
    ///     Identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document session handling.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptSmartArtTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptSmartArtTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes a PowerPoint SmartArt operation (add, manage_nodes).
    /// </summary>
    /// <param name="operation">The operation to perform: add, manage_nodes.</param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (optional, defaults to input path).</param>
    /// <param name="slideIndex">Slide index (0-based, required for all operations).</param>
    /// <param name="shapeIndex">Shape index (0-based, required for manage_nodes).</param>
    /// <param name="layout">
    ///     SmartArt layout type: BasicProcess, BasicCycle, BasicPyramid, BasicRadial, Hierarchy,
    ///     OrganizationChart, etc.
    /// </param>
    /// <param name="x">X position (optional, for add operation, defaults to 100).</param>
    /// <param name="y">Y position (optional, for add operation, defaults to 100).</param>
    /// <param name="width">Width (optional, for add operation, defaults to 400).</param>
    /// <param name="height">Height (optional, for add operation, defaults to 300).</param>
    /// <param name="action">Node action: add, edit, delete (required for manage_nodes operation).</param>
    /// <param name="targetPath">
    ///     Array of indices to target node as JSON (e.g., '[0]' for first node, '[0,1]' for second child
    ///     of first node).
    /// </param>
    /// <param name="text">Node text content (required for add/edit operations in manage_nodes).</param>
    /// <param name="position">Insert position for new node (0-based, optional for add action, defaults to append at end).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "ppt_smart_art")]
    [Description(@"Manage PowerPoint SmartArt. Supports 2 operations: add, manage_nodes.

Usage examples:
- Add SmartArt: ppt_smart_art(operation='add', path='presentation.pptx', slideIndex=0, layout='BasicProcess', x=100, y=100, width=400, height=300)
- Manage nodes: ppt_smart_art(operation='manage_nodes', path='presentation.pptx', slideIndex=0, shapeIndex=0, action='add', targetPath='[0]', text='New Node')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'add': Add a new SmartArt shape (required params: path, slideIndex, layout)
- 'manage_nodes': Manage SmartArt nodes (add, edit, delete) (required params: path, slideIndex, shapeIndex, action)")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (optional, defaults to input path)")]
        string? outputPath = null,
        [Description("Slide index (0-based, required for all operations)")]
        int slideIndex = 0,
        [Description("Shape index (0-based, required for manage_nodes)")]
        int? shapeIndex = null,
        [Description(
            "SmartArt layout type: BasicProcess, BasicCycle, BasicPyramid, BasicRadial, Hierarchy, OrganizationChart, HorizontalHierarchy, CircleArrowProcess, ClosedChevronProcess, StepDownProcess")]
        string? layout = null,
        [Description("X position (optional, for add operation, defaults to 100)")]
        float x = 100,
        [Description("Y position (optional, for add operation, defaults to 100)")]
        float y = 100,
        [Description("Width (optional, for add operation, defaults to 400)")]
        float width = 400,
        [Description("Height (optional, for add operation, defaults to 300)")]
        float height = 300,
        [Description("Node action: 'add', 'edit', 'delete' (required for manage_nodes operation)")]
        string? action = null,
        [Description(
            "Array of indices to target node as JSON (e.g., '[0]' for first node, '[0,1]' for second child of first node)")]
        string? targetPath = null,
        [Description("Node text content (required for add/edit operations in manage_nodes)")]
        string? text = null,
        [Description("Insert position for new node (0-based, optional for add action, defaults to append at end)")]
        int? position = null)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLower() switch
        {
            "add" => AddSmartArt(ctx, outputPath, slideIndex, layout, x, y, width, height),
            "manage_nodes" => ManageNodes(ctx, outputPath, slideIndex, shapeIndex, action, targetPath, text, position),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a SmartArt shape to a slide.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="layoutStr">The SmartArt layout type string.</param>
    /// <param name="x">The X position in points.</param>
    /// <param name="y">The Y position in points.</param>
    /// <param name="width">The width in points.</param>
    /// <param name="height">The height in points.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when layout is not provided or is invalid.</exception>
    private static string AddSmartArt(DocumentContext<Presentation> ctx, string? outputPath, int slideIndex,
        string? layoutStr, float x, float y, float width, float height)
    {
        if (string.IsNullOrEmpty(layoutStr))
            throw new ArgumentException("layout is required for add operation");

        if (!Enum.TryParse<SmartArtLayoutType>(layoutStr, true, out var layoutType))
            throw new ArgumentException(
                $"Invalid SmartArt layout: '{layoutStr}'. Valid layouts include: BasicProcess, Cycle, Hierarchy, etc.");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        slide.Shapes.AddSmartArt(x, y, width, height, layoutType);

        ctx.Save(outputPath);

        return $"SmartArt '{layoutStr}' added to slide {slideIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Manages SmartArt nodes (add, edit, delete).
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index (0-based).</param>
    /// <param name="action">The node action (add, edit, delete).</param>
    /// <param name="targetPathJson">JSON array of indices to target node.</param>
    /// <param name="text">The node text content.</param>
    /// <param name="position">The insert position for new node.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when required parameters are missing, shape is not SmartArt, or node path is
    ///     invalid.
    /// </exception>
    private static string ManageNodes(DocumentContext<Presentation> ctx, string? outputPath, int slideIndex,
        int? shapeIndex, string? action, string? targetPathJson, string? text, int? position)
    {
        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for manage_nodes operation");
        if (string.IsNullOrEmpty(action))
            throw new ArgumentException("action is required for manage_nodes operation");
        if (string.IsNullOrEmpty(targetPathJson))
            throw new ArgumentException("targetPath is required for manage_nodes operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        PowerPointHelper.ValidateShapeIndex(shapeIndex.Value, slide);
        var shape = slide.Shapes[shapeIndex.Value];

        if (shape is not ISmartArt smartArt)
            throw new ArgumentException(
                $"Shape at index {shapeIndex.Value} is not a SmartArt shape. It is a {shape.GetType().Name}.");

        var targetPathArray = JsonSerializer.Deserialize<int[]>(targetPathJson);
        if (targetPathArray == null || targetPathArray.Length == 0)
            throw new ArgumentException("targetPath must contain at least one index.");

        var rootIndex = targetPathArray[0];
        if (rootIndex < 0 || rootIndex >= smartArt.AllNodes.Count)
            throw new ArgumentException(
                $"Root index {rootIndex} is out of range (SmartArt has {smartArt.AllNodes.Count} root nodes).");

        ISmartArtNode targetNode;
        var currentNode = smartArt.AllNodes[rootIndex];

        if (targetPathArray.Length == 1)
        {
            targetNode = currentNode;
        }
        else
        {
            for (var i = 1; i < targetPathArray.Length; i++)
            {
                var childIndex = targetPathArray[i];
                if (childIndex < 0 || childIndex >= currentNode.ChildNodes.Count)
                    throw new ArgumentException(
                        $"Child index {childIndex} is out of range (node has {currentNode.ChildNodes.Count} children).");

                currentNode = currentNode.ChildNodes[childIndex];
            }

            targetNode = currentNode;
        }

        var sb = new StringBuilder();

        switch (action.ToLower())
        {
            case "add":
                if (string.IsNullOrEmpty(text))
                    throw new ArgumentException("text parameter is required for 'add' action.");

                ISmartArtNode newNode;
                if (position.HasValue)
                {
                    var childCount = targetNode.ChildNodes.Count;
                    if (position.Value < 0 || position.Value > childCount)
                        throw new ArgumentException(
                            $"Position {position.Value} is out of range (valid: 0-{childCount}).");

                    newNode = targetNode.ChildNodes.AddNodeByPosition(position.Value);
                    sb.AppendLine($"Node added at position {position.Value}.");
                }
                else
                {
                    newNode = targetNode.ChildNodes.AddNode();
                    sb.AppendLine("Node added at end.");
                }

                newNode.TextFrame.Text = text;
                sb.AppendLine($"Target path: {targetPathJson}");
                sb.AppendLine($"New node text: {text}");
                break;

            case "edit":
                if (string.IsNullOrEmpty(text))
                    throw new ArgumentException("text parameter is required for 'edit' action");

                targetNode.TextFrame.Text = text;
                sb.AppendLine("Node edited successfully");
                sb.AppendLine($"Target path: {targetPathJson}");
                sb.AppendLine($"New text: {text}");
                break;

            case "delete":
                var nodeText = targetNode.TextFrame?.Text ?? "(empty)";

                if (targetPathArray.Length == 1)
                    smartArt.AllNodes.RemoveNode(rootIndex);
                else
                    targetNode.Remove();

                sb.AppendLine("Node deleted successfully.");
                sb.AppendLine($"Deleted node path: {targetPathJson}");
                sb.AppendLine($"Deleted node text: {nodeText}");
                break;

            default:
                throw new ArgumentException($"Unknown action: {action}. Valid actions: add, edit, delete");
        }

        ctx.Save(outputPath);

        sb.Append(ctx.GetOutputMessage(outputPath));
        return sb.ToString();
    }
}