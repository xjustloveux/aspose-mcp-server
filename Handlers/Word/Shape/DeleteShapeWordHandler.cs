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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var shapeIndex = parameters.GetOptional<int?>("shapeIndex");

        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for delete operation");

        var doc = context.Document;
        var shapes = WordShapeHelper.GetAllShapes(doc);

        if (shapeIndex.Value < 0 || shapeIndex.Value >= shapes.Count)
            throw new ArgumentException(
                $"Shape index {shapeIndex.Value} is out of range. Document has {shapes.Count} shapes.");

        shapes[shapeIndex.Value].Remove();

        MarkModified(context);

        return $"Successfully deleted shape at index {shapeIndex.Value}.";
    }
}
