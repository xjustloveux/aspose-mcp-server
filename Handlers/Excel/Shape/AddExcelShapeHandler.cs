using Aspose.Cells;
using Aspose.Cells.Drawing;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Shape;

/// <summary>
///     Handler for adding an auto shape to an Excel worksheet.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddExcelShapeHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds an auto shape to a worksheet.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: shapeType
    ///     Optional: sheetIndex, upperLeftRow, upperLeftColumn, width, height, text
    /// </param>
    /// <returns>Success message with shape details.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractParameters(parameters);

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);

            var autoShapeType = ResolveAutoShapeType(p.ShapeType);
            var shape = worksheet.Shapes.AddAutoShape(autoShapeType,
                p.UpperLeftRow, 0, p.UpperLeftColumn, 0, p.Height, p.Width);

            if (!string.IsNullOrEmpty(p.Text))
                shape.Text = p.Text;

            MarkModified(context);

            return new SuccessResult
            {
                Message =
                    $"Shape '{p.ShapeType}' added at row {p.UpperLeftRow}, column {p.UpperLeftColumn} in sheet {p.SheetIndex}."
            };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Failed to add shape: {ex.Message}");
        }
    }

    /// <summary>
    ///     Resolves a shape type string to an AutoShapeType enum value.
    /// </summary>
    /// <param name="shapeType">The shape type string.</param>
    /// <returns>The corresponding AutoShapeType value.</returns>
    /// <exception cref="ArgumentException">Thrown when the shape type is unknown.</exception>
    internal static AutoShapeType ResolveAutoShapeType(string shapeType)
    {
        if (Enum.TryParse<AutoShapeType>(shapeType, true, out var result))
            return result;

        throw new ArgumentException(
            $"Unknown shape type: '{shapeType}'. Common types: Rectangle, RoundedRectangle, Oval, Diamond, " +
            "IsoscelesTriangle, RightTriangle, Parallelogram, Hexagon, Octagon, Star5, Star6, RightArrow, Cube, Can.");
    }

    private static AddParameters ExtractParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var shapeType = parameters.GetOptional<string?>("shapeType");
        var upperLeftRow = parameters.GetOptional("upperLeftRow", 0);
        var upperLeftColumn = parameters.GetOptional("upperLeftColumn", 0);
        var width = parameters.GetOptional("width", 100);
        var height = parameters.GetOptional("height", 100);
        var text = parameters.GetOptional<string?>("text");

        if (string.IsNullOrEmpty(shapeType))
            throw new ArgumentException("shapeType is required for add operation");

        return new AddParameters(sheetIndex, shapeType, upperLeftRow, upperLeftColumn, width, height, text);
    }

    private sealed record AddParameters(
        int SheetIndex,
        string ShapeType,
        int UpperLeftRow,
        int UpperLeftColumn,
        int Width,
        int Height,
        string? Text);
}
