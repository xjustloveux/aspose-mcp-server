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
        var slideIndex = parameters.GetRequired<int>("slideIndex");
        var shapeIndex = parameters.GetRequired<int>("shapeIndex");
        var action = parameters.GetRequired<string>("action");
        var targetPathJson = parameters.GetRequired<string>("targetPath");
        var text = parameters.GetOptional<string?>("text");
        var position = parameters.GetOptional<int?>("position");

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        PowerPointHelper.ValidateShapeIndex(shapeIndex, slide);
        var shape = slide.Shapes[shapeIndex];

        if (shape is not ISmartArt smartArt)
            throw new ArgumentException(
                $"Shape at index {shapeIndex} is not a SmartArt shape. It is a {shape.GetType().Name}.");

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

        MarkModified(context);

        return sb.ToString().TrimEnd();
    }
}
