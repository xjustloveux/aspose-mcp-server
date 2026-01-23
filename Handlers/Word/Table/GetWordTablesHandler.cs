using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.Word;
using AsposeMcpServer.Results.Word.Table;

namespace AsposeMcpServer.Handlers.Word.Table;

/// <summary>
///     Handler for getting tables from Word documents.
/// </summary>
[ResultType(typeof(GetTablesWordResult))]
public class GetWordTablesHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all tables from the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: sectionIndex.
    /// </param>
    /// <returns>JSON string containing table information.</returns>
    /// <exception cref="ArgumentException">Thrown when sectionIndex is out of range.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractGetWordTablesParameters(parameters);

        var doc = context.Document;
        var tables = WordTableHelper.GetTables(doc, p.SectionIndex);

        List<WordTableInfo> tableList = [];
        for (var i = 0; i < tables.Count; i++)
        {
            var table = tables[i];
            var rowCount = table.Rows.Count;
            var colCount = rowCount > 0 ? table.Rows[0].Cells.Count : 0;
            var precedingText = WordTableHelper.GetPrecedingText(table, 50);

            tableList.Add(new WordTableInfo
            {
                Index = i,
                Rows = rowCount,
                Columns = colCount,
                PrecedingText = !string.IsNullOrEmpty(precedingText) ? precedingText : null
            });
        }

        var result = new GetTablesWordResult
        {
            Count = tables.Count,
            SectionIndex = p.SectionIndex,
            Tables = tableList
        };

        return result;
    }

    private static GetWordTablesParameters ExtractGetWordTablesParameters(OperationParameters parameters)
    {
        var sectionIndex = parameters.GetOptional<int?>("sectionIndex");

        return new GetWordTablesParameters(sectionIndex);
    }

    private sealed record GetWordTablesParameters(int? SectionIndex);
}
