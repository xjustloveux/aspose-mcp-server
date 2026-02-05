using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Shape;

/// <summary>
///     Handler for editing a shape in an Excel worksheet.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class EditExcelShapeHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits a shape's properties.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: shapeIndex
    ///     Optional: sheetIndex, text, name, width, height, upperLeftRow, upperLeftColumn
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    public override object Execute(OperationContext<Workbook> context, OperationParameters parameters)
    {
        var p = ExtractParameters(parameters);

        try
        {
            var workbook = context.Document;
            var worksheet = ExcelHelper.GetWorksheet(workbook, p.SheetIndex);

            GetExcelShapesHandler.ValidateShapeIndex(worksheet, p.ShapeIndex);

            var shape = worksheet.Shapes[p.ShapeIndex];
            var changes = new List<string>();

            if (p.Text != null)
            {
                shape.Text = p.Text;
                changes.Add("text");
            }

            if (p.Name != null)
            {
                shape.Name = p.Name;
                changes.Add("name");
            }

            if (p.Width.HasValue)
            {
                shape.Width = p.Width.Value;
                changes.Add("width");
            }

            if (p.Height.HasValue)
            {
                shape.Height = p.Height.Value;
                changes.Add("height");
            }

            if (p.UpperLeftRow.HasValue)
            {
                shape.UpperLeftRow = p.UpperLeftRow.Value;
                changes.Add("upperLeftRow");
            }

            if (p.UpperLeftColumn.HasValue)
            {
                shape.UpperLeftColumn = p.UpperLeftColumn.Value;
                changes.Add("upperLeftColumn");
            }

            if (changes.Count == 0)
                return new SuccessResult { Message = "No changes specified." };

            MarkModified(context);

            return new SuccessResult
            {
                Message =
                    $"Shape at index {p.ShapeIndex} updated: {string.Join(", ", changes)} in sheet {p.SheetIndex}."
            };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Failed to edit shape: {ex.Message}");
        }
    }

    private static EditParameters ExtractParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var shapeIndex = parameters.GetOptional<int?>("shapeIndex");
        var text = parameters.GetOptional<string?>("text");
        var name = parameters.GetOptional<string?>("name");
        var width = parameters.GetOptional<int?>("width");
        var height = parameters.GetOptional<int?>("height");
        var upperLeftRow = parameters.GetOptional<int?>("upperLeftRow");
        var upperLeftColumn = parameters.GetOptional<int?>("upperLeftColumn");

        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for edit operation");

        return new EditParameters(sheetIndex, shapeIndex.Value, text, name, width, height, upperLeftRow,
            upperLeftColumn);
    }

    private sealed record EditParameters(
        int SheetIndex,
        int ShapeIndex,
        string? Text,
        string? Name,
        int? Width,
        int? Height,
        int? UpperLeftRow,
        int? UpperLeftColumn);
}
