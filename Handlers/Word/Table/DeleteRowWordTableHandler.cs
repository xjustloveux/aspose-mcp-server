using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
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
        var actualSectionIndex = p.SectionIndex ?? 0;
        if (actualSectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"Section index {actualSectionIndex} out of range");

        var section = doc.Sections[actualSectionIndex];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Aspose.Words.Tables.Table>().ToList();
        if (p.TableIndex < 0 || p.TableIndex >= tables.Count)
            throw new ArgumentException($"Table index {p.TableIndex} out of range");

        var table = tables[p.TableIndex];
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
