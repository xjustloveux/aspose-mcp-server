using Aspose.Cells;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Excel;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Excel.Shape;

/// <summary>
///     Handler for adding a text box to an Excel worksheet.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddTextBoxExcelShapeHandler : OperationHandlerBase<Workbook>
{
    /// <inheritdoc />
    public override string Operation => "add_textbox";

    /// <summary>
    ///     Adds a text box to a worksheet.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: text
    ///     Optional: sheetIndex, upperLeftRow, upperLeftColumn, width, height
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

            var textBox = worksheet.Shapes.AddTextBox(
                p.UpperLeftRow, 0, p.UpperLeftColumn, 0, p.Height, p.Width);

            textBox.Text = p.Text;

            MarkModified(context);

            return new SuccessResult
            {
                Message =
                    $"TextBox added at row {p.UpperLeftRow}, column {p.UpperLeftColumn} in sheet {p.SheetIndex} with text '{p.Text}'."
            };
        }
        catch (CellsException ex)
        {
            throw new ArgumentException($"Failed to add textbox: {ex.Message}");
        }
    }

    private static AddTextBoxParameters ExtractParameters(OperationParameters parameters)
    {
        var sheetIndex = parameters.GetOptional("sheetIndex", 0);
        var text = parameters.GetOptional<string?>("text");
        var upperLeftRow = parameters.GetOptional("upperLeftRow", 0);
        var upperLeftColumn = parameters.GetOptional("upperLeftColumn", 0);
        var width = parameters.GetOptional("width", 200);
        var height = parameters.GetOptional("height", 50);

        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("text is required for add_textbox operation");

        return new AddTextBoxParameters(sheetIndex, text, upperLeftRow, upperLeftColumn, width, height);
    }

    private sealed record AddTextBoxParameters(
        int SheetIndex,
        string Text,
        int UpperLeftRow,
        int UpperLeftColumn,
        int Width,
        int Height);
}
