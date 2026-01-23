using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Table;

/// <summary>
///     Handler for deleting tables from Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractDeleteWordTableParameters(parameters);

        var doc = context.Document;
        var tables = WordTableHelper.GetTables(doc, p.SectionIndex);

        if (p.TableIndex < 0 || p.TableIndex >= tables.Count)
            throw new ArgumentException($"Table index {p.TableIndex} out of range");

        tables[p.TableIndex].Remove();

        MarkModified(context);

        return new SuccessResult
            { Message = $"Successfully deleted table #{p.TableIndex}. Remaining tables: {tables.Count - 1}." };
    }

    private static DeleteWordTableParameters ExtractDeleteWordTableParameters(OperationParameters parameters)
    {
        var tableIndex = parameters.GetOptional("tableIndex", 0);
        var sectionIndex = parameters.GetOptional<int?>("sectionIndex");

        return new DeleteWordTableParameters(tableIndex, sectionIndex);
    }

    private sealed record DeleteWordTableParameters(int TableIndex, int? SectionIndex);
}
