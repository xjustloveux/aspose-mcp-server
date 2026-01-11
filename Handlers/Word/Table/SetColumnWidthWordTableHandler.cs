using Aspose.Words;
using Aspose.Words.Tables;
using AsposeMcpServer.Core.Handlers;
using WordTable = Aspose.Words.Tables.Table;

namespace AsposeMcpServer.Handlers.Word.Table;

/// <summary>
///     Handler for setting column width in Word document tables.
/// </summary>
public class SetColumnWidthWordTableHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "set_column_width";

    /// <summary>
    ///     Sets column width for a table.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: columnIndex, columnWidth.
    ///     Optional: tableIndex (default 0), sectionIndex.
    /// </param>
    /// <returns>Success message with updated cell count.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or indices are out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when the table has no rows.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var columnIndex = parameters.GetOptional<int?>("columnIndex");
        var columnWidth = parameters.GetOptional<double?>("columnWidth");

        if (!columnIndex.HasValue)
            throw new ArgumentException("columnIndex is required for set_column_width operation");
        if (!columnWidth.HasValue)
            throw new ArgumentException("columnWidth is required for set_column_width operation");
        if (columnWidth.Value <= 0)
            throw new ArgumentException($"Column width {columnWidth.Value} must be greater than 0");

        var tableIndex = parameters.GetOptional("tableIndex", 0);
        var sectionIndex = parameters.GetOptional<int?>("sectionIndex");

        var doc = context.Document;
        var actualSectionIndex = sectionIndex ?? 0;
        if (actualSectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"Section index {actualSectionIndex} out of range");

        var section = doc.Sections[actualSectionIndex];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<WordTable>().ToList();
        if (tableIndex < 0 || tableIndex >= tables.Count)
            throw new ArgumentException($"Table index {tableIndex} out of range");

        var table = tables[tableIndex];
        if (table.Rows.Count == 0)
            throw new InvalidOperationException($"Table {tableIndex} has no rows");

        var firstRow = table.Rows[0];
        if (columnIndex.Value < 0 || columnIndex.Value >= firstRow.Cells.Count)
            throw new ArgumentException($"Column index {columnIndex.Value} out of range");

        var cellsUpdated = 0;
        foreach (var row in table.Rows.Cast<Row>())
            if (columnIndex.Value < row.Cells.Count)
            {
                row.Cells[columnIndex.Value].CellFormat.PreferredWidth = PreferredWidth.FromPoints(columnWidth.Value);
                cellsUpdated++;
            }

        MarkModified(context);

        return Success(
            $"Successfully set column {columnIndex.Value} width to {columnWidth.Value} pt ({cellsUpdated} cells updated).");
    }
}
