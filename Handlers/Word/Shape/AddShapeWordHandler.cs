using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Shape;

/// <summary>
///     Handler for adding generic shapes to Word documents.
/// </summary>
public class AddShapeWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a generic shape to the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: shapeType, width, height
    ///     Optional: x, y
    /// </param>
    /// <returns>Success message with shape details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var shapeType = parameters.GetOptional<string?>("shapeType");
        var width = parameters.GetOptional<double?>("width");
        var height = parameters.GetOptional<double?>("height");
        var x = parameters.GetOptional("x", 100.0);
        var y = parameters.GetOptional("y", 100.0);

        if (string.IsNullOrEmpty(shapeType))
            throw new ArgumentException("shapeType is required for add operation");
        if (!width.HasValue)
            throw new ArgumentException("width is required for add operation");
        if (!height.HasValue)
            throw new ArgumentException("height is required for add operation");

        var shapeTypeEnum = WordShapeHelper.ParseShapeType(shapeType);

        var doc = context.Document;
        var builder = new DocumentBuilder(doc);
        var shape = builder.InsertShape(shapeTypeEnum, width.Value, height.Value);
        shape.Left = x;
        shape.Top = y;

        MarkModified(context);

        return $"Successfully added {shapeType} shape.";
    }
}
