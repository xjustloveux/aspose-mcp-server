using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Excel.Shape;

namespace AsposeMcpServer.Handlers.Excel.Shape;

/// <summary>
///     Handler for getting shapes from an Excel worksheet.
/// </summary>
[ResultType(typeof(GetShapesExcelResult))]
public class GetExcelShapesHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all shapes from a worksheet.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: sheetIndex (default: 0), shapeIndex (specific shape)
    /// </param>
    /// <returns>Shape information result.</returns>
    /// <exception cref="ArgumentException">Thrown when parameters are invalid.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var shapeIndex = parameters.GetOptional<int?>("shapeIndex");

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            var shapes = new List<ExcelShapeInfo>();

            if (shapeIndex.HasValue)
            {
                ValidateShapeIndex(worksheet, shapeIndex.Value);
                shapes.Add(BuildShapeInfo(worksheet.Shapes[shapeIndex.Value], shapeIndex.Value));
            }
            else
            {
                for (var i = 0; i < worksheet.Shapes.Count; i++)
                    shapes.Add(BuildShapeInfo(worksheet.Shapes[i], i));
            }

            return new GetShapesExcelResult
            {
                Count = shapes.Count,
                SheetIndex = sheetIndex,
                Items = shapes,
                Message = shapes.Count == 0 ? "No shapes found in the worksheet." : null
            };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Failed to get shapes from sheet {sheetIndex}: {ex.Message}");
        }
    }

    /// <summary>
    ///     Builds shape information from a Shape object.
    /// </summary>
    /// <param name="shape">The shape.</param>
    /// <param name="index">The shape index.</param>
    /// <returns>Shape information record.</returns>
    internal static ExcelShapeInfo BuildShapeInfo(Aspose.Cells.Drawing.Shape shape, int index)
    {
        return new ExcelShapeInfo
        {
            Index = index,
            Name = shape.Name ?? $"Shape{index}",
            Type = shape.MsoDrawingType.ToString(),
            Text = string.IsNullOrEmpty(shape.Text) ? null : shape.Text,
            UpperLeftRow = shape.UpperLeftRow,
            UpperLeftColumn = shape.UpperLeftColumn,
            Width = shape.Width,
            Height = shape.Height
        };
    }

    /// <summary>
    ///     Validates a shape index.
    /// </summary>
    /// <param name="worksheet">The worksheet.</param>
    /// <param name="shapeIndex">The shape index.</param>
    /// <exception cref="ArgumentException">Thrown when the index is out of range.</exception>
    internal static void ValidateShapeIndex(Worksheet worksheet, int shapeIndex)
    {
        if (shapeIndex < 0 || shapeIndex >= worksheet.Shapes.Count)
            throw new ArgumentException(
                $"Shape index {shapeIndex} is out of range (worksheet has {worksheet.Shapes.Count} shapes)");
    }
}
