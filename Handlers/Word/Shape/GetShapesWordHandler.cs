using System.Text;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Shape;

/// <summary>
///     Handler for getting all shapes from Word documents.
/// </summary>
public class GetShapesWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all shapes from the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">No parameters required.</param>
    /// <returns>Formatted string containing shape information.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var doc = context.Document;
        var shapes = WordShapeHelper.GetAllShapes(doc);

        var result = new StringBuilder();
        result.AppendLine("=== Document Shapes ===\n");
        result.AppendLine($"Total Shapes: {shapes.Count}\n");

        if (shapes.Count == 0)
        {
            result.AppendLine("No shapes found");
            return result.ToString();
        }

        for (var i = 0; i < shapes.Count; i++)
        {
            var shape = shapes[i];
            result.AppendLine($"Shape {i}:");
            result.AppendLine($"  Type: {shape.ShapeType}");
            result.AppendLine($"  Name: {shape.Name ?? "(No name)"}");
            result.AppendLine($"  Size: {shape.Width} x {shape.Height} pt");
            result.AppendLine($"  Position: X={shape.Left}, Y={shape.Top}");
            result.AppendLine();
        }

        return result.ToString();
    }
}
