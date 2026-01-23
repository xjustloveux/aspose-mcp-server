using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Shape;

/// <summary>
///     Handler for adding generic shapes to Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    /// <exception cref="ArgumentException">Thrown when required parameters are missing.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractAddShapeParameters(parameters);

        var shapeTypeEnum = WordShapeHelper.ParseShapeType(p.ShapeType);

        var doc = context.Document;
        var builder = new DocumentBuilder(doc);
        var shape = builder.InsertShape(shapeTypeEnum, p.Width, p.Height);
        shape.Left = p.X;
        shape.Top = p.Y;

        MarkModified(context);

        return new SuccessResult { Message = $"Successfully added {p.ShapeType} shape." };
    }

    private static AddShapeParameters ExtractAddShapeParameters(OperationParameters parameters)
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

        return new AddShapeParameters(shapeType, width.Value, height.Value, x, y);
    }

    private sealed record AddShapeParameters(
        string ShapeType,
        double Width,
        double Height,
        double X,
        double Y);
}
