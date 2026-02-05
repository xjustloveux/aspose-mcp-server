using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Shape;

/// <summary>
///     Handler for deleting a shape from an Excel worksheet.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class DeleteExcelShapeHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a shape from the worksheet.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: shapeIndex
    ///     Optional: sheetIndex (default: 0)
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var shapeIndex = parameters.GetOptional<int?>("shapeIndex");

        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for delete operation");

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            GetExcelShapesHandler.ValidateShapeIndex(worksheet, shapeIndex.Value);

            var shapeName = worksheet.Shapes[shapeIndex.Value].Name ?? $"Shape{shapeIndex.Value}";
            worksheet.Shapes.RemoveAt(shapeIndex.Value);

            MarkModified(context);

            return new SuccessResult
            {
                Message = $"Shape '{shapeName}' (index {shapeIndex.Value}) deleted from sheet {sheetIndex}."
            };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Failed to delete shape: {ex.Message}");
        }
    }
}
