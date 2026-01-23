using Aspose.Words;
using Aspose.Words.Tables;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Table;

/// <summary>
///     Handler for deleting columns from Word document tables.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractDeleteColumnParameters(parameters);

        var doc = context.Document;
        var tables = WordTableHelper.GetTables(doc, p.SectionIndex);

        if (p.TableIndex < 0 || p.TableIndex >= tables.Count)
            throw new ArgumentException($"Table index {p.TableIndex} out of range");

        var table = tables[p.TableIndex];
        if (table.Rows.Count == 0)
            throw new InvalidOperationException($"Table {p.TableIndex} has no rows");

        var firstRow = table.Rows[0];
        if (p.ColumnIndex < 0 || p.ColumnIndex >= firstRow.Cells.Count)
            throw new ArgumentException($"Column index {p.ColumnIndex} out of range");

        var deletedCount = 0;
        foreach (var row in table.Rows.Cast<Row>()) // NOSONAR S3267 - Loop modifies collection
            if (p.ColumnIndex < row.Cells.Count)
            {
                row.Cells[p.ColumnIndex].Remove();
                deletedCount++;
            }

        MarkModified(context);

        return new SuccessResult
            { Message = $"Successfully deleted column #{p.ColumnIndex} ({deletedCount} cells removed)." };
    }

    private static DeleteColumnParameters ExtractDeleteColumnParameters(OperationParameters parameters)
    {
        var columnIndex = parameters.GetOptional<int?>("columnIndex");
        if (!columnIndex.HasValue)
            throw new ArgumentException("columnIndex is required for delete_column operation");

        var tableIndex = parameters.GetOptional("tableIndex", 0);
        var sectionIndex = parameters.GetOptional<int?>("sectionIndex");

        return new DeleteColumnParameters(columnIndex.Value, tableIndex, sectionIndex);
    }

    private sealed record DeleteColumnParameters(int ColumnIndex, int TableIndex, int? SectionIndex);
}
