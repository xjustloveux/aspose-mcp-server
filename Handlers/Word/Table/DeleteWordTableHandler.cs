using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Table;

/// <summary>
///     Handler for deleting tables from Word documents.
/// </summary>
public class DeleteWordTableHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a table from the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: tableIndex (default 0), sectionIndex.
    /// </param>
    /// <returns>Success message with remaining table count.</returns>
    /// <exception cref="ArgumentException">Thrown when tableIndex or sectionIndex is out of range.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var tableIndex = parameters.GetOptional("tableIndex", 0);
        var sectionIndex = parameters.GetOptional<int?>("sectionIndex");

        var doc = context.Document;
        var tables = WordTableHelper.GetTables(doc, sectionIndex);

        if (tableIndex < 0 || tableIndex >= tables.Count)
            throw new ArgumentException($"Table index {tableIndex} out of range");

        tables[tableIndex].Remove();

        MarkModified(context);

        return Success($"Successfully deleted table #{tableIndex}. Remaining tables: {tables.Count - 1}.");
    }
}
