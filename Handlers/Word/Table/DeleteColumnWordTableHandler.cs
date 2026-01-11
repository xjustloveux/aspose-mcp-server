using Aspose.Words;
using Aspose.Words.Tables;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Table;

/// <summary>
///     Handler for deleting columns from Word document tables.
/// </summary>
public class DeleteColumnWordTableHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete_column";

    /// <summary>
    ///     Deletes a column from a table.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: columnIndex.
    ///     Optional: tableIndex (default 0), sectionIndex.
    /// </param>
    /// <returns>Success message with deleted cell count.</returns>
    /// <exception cref="ArgumentException">Thrown when columnIndex is missing or indices are out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when the table has no rows.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var columnIndex = parameters.GetOptional<int?>("columnIndex");
        if (!columnIndex.HasValue)
            throw new ArgumentException("columnIndex is required for delete_column operation");

        var tableIndex = parameters.GetOptional("tableIndex", 0);
        var sectionIndex = parameters.GetOptional<int?>("sectionIndex");

        var doc = context.Document;
        var tables = WordTableHelper.GetTables(doc, sectionIndex);

        if (tableIndex < 0 || tableIndex >= tables.Count)
            throw new ArgumentException($"Table index {tableIndex} out of range");

        var table = tables[tableIndex];
        if (table.Rows.Count == 0)
            throw new InvalidOperationException($"Table {tableIndex} has no rows");

        var firstRow = table.Rows[0];
        if (columnIndex.Value < 0 || columnIndex.Value >= firstRow.Cells.Count)
            throw new ArgumentException($"Column index {columnIndex.Value} out of range");

        var deletedCount = 0;
        foreach (var row in table.Rows.Cast<Row>())
            if (columnIndex.Value < row.Cells.Count)
            {
                row.Cells[columnIndex.Value].Remove();
                deletedCount++;
            }

        MarkModified(context);

        return Success($"Successfully deleted column #{columnIndex.Value} ({deletedCount} cells removed).");
    }
}
