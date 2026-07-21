using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Table;

/// <summary>
///     Handler for deleting rows from Word document tables.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractDeleteRowParameters(parameters);

        var doc = context.Document;
        // Omitted sectionIndex means the document-wide flat table index, matching 'get'.
        var table = WordTableHelper.GetTable(doc, p.TableIndex, p.SectionIndex);
        if (p.RowIndex < 0 || p.RowIndex >= table.Rows.Count)
            throw new ArgumentException($"Row index {p.RowIndex} out of range");

        var rowToDelete = table.Rows[p.RowIndex];
        rowToDelete.Remove();

        MarkModified(context);

        return new SuccessResult
            { Message = $"Successfully deleted row #{p.RowIndex}. Remaining rows: {table.Rows.Count}." };
    }

    private static DeleteRowParameters ExtractDeleteRowParameters(OperationParameters parameters)
    {
        var rowIndex = parameters.GetOptional<int?>("rowIndex");
        if (!rowIndex.HasValue)
            throw new ArgumentException("rowIndex is required for delete_row operation");

        var tableIndex = parameters.GetOptional("tableIndex", 0);
        var sectionIndex = parameters.GetOptional<int?>("sectionIndex");

        return new DeleteRowParameters(rowIndex.Value, tableIndex, sectionIndex);
    }

    private sealed record DeleteRowParameters(int RowIndex, int TableIndex, int? SectionIndex);
}
