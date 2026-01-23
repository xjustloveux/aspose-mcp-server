using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Table;

/// <summary>
///     Handler for editing cell content in PowerPoint tables.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class EditPptTableCellHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "edit_cell";

    /// <summary>
    ///     Edits a cell in a table.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: shapeIndex, rowIndex, columnIndex, text.
    ///     Optional: slideIndex.
    /// </param>
    /// <returns>Success message with update details.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var cellParams = ExtractEditCellParameters(parameters);

        var presentation = context.Document;
        var slide = PptTableHelper.GetSlide(presentation, cellParams.SlideIndex);
        var table = PptTableHelper.GetTable(slide, cellParams.ShapeIndex);

        PptTableHelper.ValidateRowIndex(table, cellParams.RowIndex);
        PptTableHelper.ValidateColumnIndex(table, cellParams.ColumnIndex);

        table[cellParams.RowIndex, cellParams.ColumnIndex].TextFrame.Text = cellParams.Text;

        MarkModified(context);

        return new SuccessResult { Message = $"Cell [{cellParams.RowIndex},{cellParams.ColumnIndex}] updated." };
    }

    /// <summary>
    ///     Extracts edit cell parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted edit cell parameters.</returns>
    private static EditCellParameters ExtractEditCellParameters(OperationParameters parameters)
    {
        return new EditCellParameters(
            parameters.GetOptional("slideIndex", 0),
            parameters.GetRequired<int>("shapeIndex"),
            parameters.GetRequired<int>("rowIndex"),
            parameters.GetRequired<int>("columnIndex"),
            parameters.GetRequired<string>("text")
        );
    }

    /// <summary>
    ///     Record for holding edit cell parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ShapeIndex">The shape index.</param>
    /// <param name="RowIndex">The row index.</param>
    /// <param name="ColumnIndex">The column index.</param>
    /// <param name="Text">The text to set.</param>
    private sealed record EditCellParameters(
        int SlideIndex,
        int ShapeIndex,
        int RowIndex,
        int ColumnIndex,
        string Text);
}
