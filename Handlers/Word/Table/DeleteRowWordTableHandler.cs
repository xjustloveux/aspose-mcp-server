using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Table;

/// <summary>
///     Handler for deleting rows from Word document tables.
/// </summary>
public class DeleteRowWordTableHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete_row";

    /// <summary>
    ///     Deletes a row from a table.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: rowIndex.
    ///     Optional: tableIndex (default 0), sectionIndex.
    /// </param>
    /// <returns>Success message with remaining row count.</returns>
    /// <exception cref="ArgumentException">Thrown when rowIndex is missing or indices are out of range.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var rowIndex = parameters.GetOptional<int?>("rowIndex");
        if (!rowIndex.HasValue)
            throw new ArgumentException("rowIndex is required for delete_row operation");

        var tableIndex = parameters.GetOptional("tableIndex", 0);
        var sectionIndex = parameters.GetOptional<int?>("sectionIndex");

        var doc = context.Document;
        var actualSectionIndex = sectionIndex ?? 0;
        if (actualSectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"Section index {actualSectionIndex} out of range");

        var section = doc.Sections[actualSectionIndex];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Aspose.Words.Tables.Table>().ToList();
        if (tableIndex < 0 || tableIndex >= tables.Count)
            throw new ArgumentException($"Table index {tableIndex} out of range");

        var table = tables[tableIndex];
        if (rowIndex.Value < 0 || rowIndex.Value >= table.Rows.Count)
            throw new ArgumentException($"Row index {rowIndex.Value} out of range");

        var rowToDelete = table.Rows[rowIndex.Value];
        rowToDelete.Remove();

        MarkModified(context);

        return Success($"Successfully deleted row #{rowIndex.Value}. Remaining rows: {table.Rows.Count}.");
    }
}
