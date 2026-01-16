using System.Text;
using System.Text.Json;
using Aspose.Slides;
using Aspose.Slides.SmartArt;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.SmartArt;

/// <summary>
///     Handler for managing SmartArt nodes (add, edit, delete).
/// </summary>
public class ManageSmartArtNodesHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "manage_nodes";

    /// <summary>
    ///     Manages SmartArt nodes.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex, shapeIndex, action, targetPath.
    ///     Optional: text (required for add/edit), position.
    /// </param>
    /// <returns>Success message with node operation details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var nodeParams = ExtractNodeParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, nodeParams.SlideIndex);
        var smartArt = GetSmartArtShape(slide, nodeParams.ShapeIndex);
        var targetPathArray = ParseTargetPath(nodeParams.TargetPath);
        var (targetNode, rootIndex) = NavigateToTargetNode(smartArt, targetPathArray);

        var actionContext = new ActionContext(smartArt, targetNode, targetPathArray, rootIndex);
        var result = ExecuteAction(nodeParams, actionContext);

        MarkModified(context);
        return result;
    }

    /// <summary>
    ///     Gets the SmartArt shape from a slide.
    /// </summary>
    /// <param name="slide">The slide containing the shape.</param>
    /// <param name="shapeIndex">The index of the shape.</param>
    /// <returns>The SmartArt shape.</returns>
    /// <exception cref="ArgumentException">Thrown when the shape at the specified index is not a SmartArt shape.</exception>
    private static ISmartArt GetSmartArtShape(ISlide slide, int shapeIndex)
    {
        PowerPointHelper.ValidateShapeIndex(shapeIndex, slide);
        var shape = slide.Shapes[shapeIndex];
        if (shape is not ISmartArt smartArt)
            throw new ArgumentException(
                $"Shape at index {shapeIndex} is not a SmartArt shape. It is a {shape.GetType().Name}.");
        return smartArt;
    }

    /// <summary>
    ///     Parses the target path JSON string to an integer array.
    /// </summary>
    /// <param name="targetPathJson">The JSON string containing the target path.</param>
    /// <returns>The parsed target path array.</returns>
    /// <exception cref="ArgumentException">Thrown when the target path is empty or null.</exception>
    private static int[] ParseTargetPath(string targetPathJson)
    {
        var targetPathArray = JsonSerializer.Deserialize<int[]>(targetPathJson);
        if (targetPathArray == null || targetPathArray.Length == 0)
            throw new ArgumentException("targetPath must contain at least one index.");
        return targetPathArray;
    }

    /// <summary>
    ///     Navigates to the target node in the SmartArt hierarchy.
    /// </summary>
    /// <param name="smartArt">The SmartArt shape.</param>
    /// <param name="targetPath">The target path array.</param>
    /// <returns>A tuple containing the target node and root index.</returns>
    /// <exception cref="ArgumentException">Thrown when a path index is out of range.</exception>
    private static (ISmartArtNode targetNode, int rootIndex) NavigateToTargetNode(ISmartArt smartArt, int[] targetPath)
    {
        var rootIndex = targetPath[0];
        if (rootIndex < 0 || rootIndex >= smartArt.AllNodes.Count)
            throw new ArgumentException(
                $"Root index {rootIndex} is out of range (SmartArt has {smartArt.AllNodes.Count} root nodes).");

        var currentNode = smartArt.AllNodes[rootIndex];
        for (var i = 1; i < targetPath.Length; i++)
        {
            var childIndex = targetPath[i];
            if (childIndex < 0 || childIndex >= currentNode.ChildNodes.Count)
                throw new ArgumentException(
                    $"Child index {childIndex} is out of range (node has {currentNode.ChildNodes.Count} children).");
            currentNode = currentNode.ChildNodes[childIndex];
        }

        return (currentNode, rootIndex);
    }

    /// <summary>
    ///     Executes the specified action on the SmartArt node.
    /// </summary>
    /// <param name="nodeParams">The node parameters containing action, text, and position.</param>
    /// <param name="actionContext">The action context containing resolved SmartArt elements.</param>
    /// <returns>The result message.</returns>
    /// <exception cref="ArgumentException">Thrown when the action is unknown or invalid.</exception>
    private static string ExecuteAction(NodeParameters nodeParams, ActionContext actionContext)
    {
        return nodeParams.Action.ToLower() switch
        {
            "add" => ExecuteAddAction(actionContext.TargetNode, nodeParams.Text, nodeParams.Position,
                nodeParams.TargetPath),
            "edit" => ExecuteEditAction(actionContext.TargetNode, nodeParams.Text, nodeParams.TargetPath),
            "delete" => ExecuteDeleteAction(actionContext.SmartArt, actionContext.TargetNode,
                actionContext.TargetPathArray, nodeParams.TargetPath, actionContext.RootIndex),
            _ => throw new ArgumentException($"Unknown action: {nodeParams.Action}. Valid actions: add, edit, delete")
        };
    }

    /// <summary>
    ///     Executes the add action to add a new child node.
    /// </summary>
    /// <param name="targetNode">The target parent node.</param>
    /// <param name="text">The text for the new node.</param>
    /// <param name="position">The optional position for the new node.</param>
    /// <param name="targetPathJson">The target path JSON string for the result message.</param>
    /// <returns>The result message.</returns>
    /// <exception cref="ArgumentException">Thrown when text is missing or position is out of range.</exception>
    private static string ExecuteAddAction(ISmartArtNode targetNode, string? text, int? position, string targetPathJson)
    {
        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("text parameter is required for 'add' action.");

        var sb = new StringBuilder();
        ISmartArtNode newNode;

        if (position.HasValue)
        {
            var childCount = targetNode.ChildNodes.Count;
            if (position.Value < 0 || position.Value > childCount)
                throw new ArgumentException($"Position {position.Value} is out of range (valid: 0-{childCount}).");
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
        return sb.ToString().TrimEnd();
    }

    /// <summary>
    ///     Executes the edit action to update node text.
    /// </summary>
    /// <param name="targetNode">The target node to edit.</param>
    /// <param name="text">The new text for the node.</param>
    /// <param name="targetPathJson">The target path JSON string for the result message.</param>
    /// <returns>The result message.</returns>
    /// <exception cref="ArgumentException">Thrown when text is missing.</exception>
    private static string ExecuteEditAction(ISmartArtNode targetNode, string? text, string targetPathJson)
    {
        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("text parameter is required for 'edit' action");

        targetNode.TextFrame.Text = text;
        var sb = new StringBuilder();
        sb.AppendLine("Node edited successfully");
        sb.AppendLine($"Target path: {targetPathJson}");
        sb.AppendLine($"New text: {text}");
        return sb.ToString().TrimEnd();
    }

    /// <summary>
    ///     Executes the delete action to remove a node.
    /// </summary>
    /// <param name="smartArt">The SmartArt shape.</param>
    /// <param name="targetNode">The target node to delete.</param>
    /// <param name="targetPath">The target path array.</param>
    /// <param name="targetPathJson">The target path JSON string for the result message.</param>
    /// <param name="rootIndex">The root index.</param>
    /// <returns>The result message.</returns>
    private static string ExecuteDeleteAction(ISmartArt smartArt, ISmartArtNode targetNode, int[] targetPath,
        string targetPathJson, int rootIndex)
    {
        var nodeText = targetNode.TextFrame?.Text ?? "(empty)";

        if (targetPath.Length == 1)
            smartArt.AllNodes.RemoveNode(rootIndex);
        else
            targetNode.Remove();

        var sb = new StringBuilder();
        sb.AppendLine("Node deleted successfully.");
        sb.AppendLine($"Deleted node path: {targetPathJson}");
        sb.AppendLine($"Deleted node text: {nodeText}");
        return sb.ToString().TrimEnd();
    }

    /// <summary>
    ///     Extracts node parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted node parameters.</returns>
    private static NodeParameters ExtractNodeParameters(OperationParameters parameters)
    {
        return new NodeParameters(
            parameters.GetRequired<int>("slideIndex"),
            parameters.GetRequired<int>("shapeIndex"),
            parameters.GetRequired<string>("action"),
            parameters.GetRequired<string>("targetPath"),
            parameters.GetOptional<string?>("text"),
            parameters.GetOptional<int?>("position")
        );
    }

    /// <summary>
    ///     Record for holding SmartArt node management parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ShapeIndex">The shape index.</param>
    /// <param name="Action">The action to perform (add, edit, delete).</param>
    /// <param name="TargetPath">The target path JSON string.</param>
    /// <param name="Text">The optional text for the node.</param>
    /// <param name="Position">The optional position for new nodes.</param>
    private sealed record NodeParameters(
        int SlideIndex,
        int ShapeIndex,
        string Action,
        string TargetPath,
        string? Text,
        int? Position);

    /// <summary>
    ///     Record for holding resolved SmartArt action context.
    /// </summary>
    /// <param name="SmartArt">The SmartArt shape.</param>
    /// <param name="TargetNode">The target node.</param>
    /// <param name="TargetPathArray">The parsed target path array.</param>
    /// <param name="RootIndex">The root node index.</param>
    private sealed record ActionContext(
        ISmartArt SmartArt,
        ISmartArtNode TargetNode,
        int[] TargetPathArray,
        int RootIndex);
}
