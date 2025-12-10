using System.Text.Json.Nodes;
using Aspose.Pdf;

namespace AsposeMcpServer.Tools;

public class PdfEditTableTool : IAsposeTool
{
    public string Description => "Edit table cell content in PDF document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "PDF file path"
            },
            pageIndex = new
            {
                type = "number",
                description = "Page index (1-based)"
            },
            tableIndex = new
            {
                type = "number",
                description = "Table index on the page (0-based)"
            },
            row = new
            {
                type = "number",
                description = "Row index (0-based)"
            },
            column = new
            {
                type = "number",
                description = "Column index (0-based)"
            },
            text = new
            {
                type = "string",
                description = "New cell text"
            }
        },
        required = new[] { "path", "pageIndex", "tableIndex", "row", "column", "text" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");
        var tableIndex = arguments?["tableIndex"]?.GetValue<int>() ?? throw new ArgumentException("tableIndex is required");
        var row = arguments?["row"]?.GetValue<int>() ?? throw new ArgumentException("row is required");
        var column = arguments?["column"]?.GetValue<int>() ?? throw new ArgumentException("column is required");
        var text = arguments?["text"]?.GetValue<string>() ?? throw new ArgumentException("text is required");

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
        {
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");
        }

        var page = document.Pages[pageIndex];
        var tables = page.Paragraphs.OfType<Table>().ToList();
        
        if (tableIndex < 0 || tableIndex >= tables.Count)
        {
            throw new ArgumentException($"tableIndex must be between 0 and {tables.Count - 1}");
        }

        var table = tables[tableIndex];
        if (row < 0 || row >= table.Rows.Count)
        {
            throw new ArgumentException($"row must be between 0 and {table.Rows.Count - 1}");
        }

        var tableRow = table.Rows[row];
        if (column < 0 || column >= tableRow.Cells.Count)
        {
            throw new ArgumentException($"column must be between 0 and {tableRow.Cells.Count - 1}");
        }

        var cell = tableRow.Cells[column];
        cell.Paragraphs.Clear();
        cell.Paragraphs.Add(new Aspose.Pdf.Text.TextFragment(text));

        document.Save(path);
        return await Task.FromResult($"Table cell [{row}, {column}] updated on page {pageIndex}: {path}");
    }
}

