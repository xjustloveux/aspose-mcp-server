using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Shape;

/// <summary>
///     Handler for deleting shapes from Word documents.
/// </summary>
public class DeleteShapeWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a shape from the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: shapeIndex
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when shapeIndex is missing or out of range.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractDeleteShapeParameters(parameters);

        var doc = context.Document;
        var shapes = WordShapeHelper.GetAllShapes(doc);

        if (p.ShapeIndex < 0 || p.ShapeIndex >= shapes.Count)
            throw new ArgumentException(
                $"Shape index {p.ShapeIndex} is out of range. Document has {shapes.Count} shapes.");

        shapes[p.ShapeIndex].Remove();

        MarkModified(context);

        return $"Successfully deleted shape at index {p.ShapeIndex}.";
    }

    private static DeleteShapeParameters ExtractDeleteShapeParameters(OperationParameters parameters)
    {
        var shapeIndex = parameters.GetOptional<int?>("shapeIndex");

        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for delete operation");

        return new DeleteShapeParameters(shapeIndex.Value);
    }

    private sealed record DeleteShapeParameters(int ShapeIndex);
}
