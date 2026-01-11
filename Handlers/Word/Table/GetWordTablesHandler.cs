using System.Text.Json;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Table;

/// <summary>
///     Handler for getting tables from Word documents.
/// </summary>
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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var sectionIndex = parameters.GetOptional<int?>("sectionIndex");

        var doc = context.Document;
        var tables = WordTableHelper.GetTables(doc, sectionIndex);

        List<object> tableList = [];
        for (var i = 0; i < tables.Count; i++)
        {
            var table = tables[i];
            var rowCount = table.Rows.Count;
            var colCount = rowCount > 0 ? table.Rows[0].Cells.Count : 0;
            var precedingText = WordTableHelper.GetPrecedingText(table, 50);

            tableList.Add(new
            {
                index = i,
                rows = rowCount,
                columns = colCount,
                precedingText = !string.IsNullOrEmpty(precedingText) ? precedingText : null
            });
        }

        var result = new
        {
            count = tables.Count,
            sectionIndex,
            tables = tableList
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }
}
