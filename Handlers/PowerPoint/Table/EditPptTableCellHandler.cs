using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Table;

/// <summary>
///     Handler for editing cell content in PowerPoint tables.
/// </summary>
public class EditPptTableCellHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "edit_cell";

    /// <summary>
    ///     Edits a cell in a table.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: shapeIndex, rowIndex, columnIndex, text
    ///     Optional: slideIndex
    /// </param>
    /// <returns>Success message with update details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetOptional("slideIndex", 0);
        var shapeIndex = parameters.GetOptional<int?>("shapeIndex");
        var rowIndex = parameters.GetOptional<int?>("rowIndex");
        var columnIndex = parameters.GetOptional<int?>("columnIndex");
        var text = parameters.GetOptional<string?>("text");

        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for edit_cell operation");
        if (!rowIndex.HasValue)
            throw new ArgumentException("rowIndex is required for edit_cell operation");
        if (!columnIndex.HasValue)
            throw new ArgumentException("columnIndex is required for edit_cell operation");
        if (text == null)
            throw new ArgumentException("text is required for edit_cell operation");

        var presentation = context.Document;
        var slide = PptTableHelper.GetSlide(presentation, slideIndex);
        var table = PptTableHelper.GetTable(slide, shapeIndex.Value);

        PptTableHelper.ValidateRowIndex(table, rowIndex.Value);
        PptTableHelper.ValidateColumnIndex(table, columnIndex.Value);

        table[rowIndex.Value, columnIndex.Value].TextFrame.Text = text;

        MarkModified(context);

        return Success($"Cell [{rowIndex.Value},{columnIndex.Value}] updated.");
    }
}
