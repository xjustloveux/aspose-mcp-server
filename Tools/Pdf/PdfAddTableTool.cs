using System.Text.Json.Nodes;
using Aspose.Pdf;

namespace AsposeMcpServer.Tools;

public class PdfAddTableTool : IAsposeTool
{
    public string Description => "Add a table to a PDF document";

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
            rows = new
            {
                type = "number",
                description = "Number of rows"
            },
            columns = new
            {
                type = "number",
                description = "Number of columns"
            },
            data = new
            {
                type = "array",
                description = "2D array of cell data (optional)",
                items = new
                {
                    type = "array",
                    items = new { type = "string" }
                }
            }
        },
        required = new[] { "path", "pageIndex", "rows", "columns" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");
        var rows = arguments?["rows"]?.GetValue<int>() ?? throw new ArgumentException("rows is required");
        var columns = arguments?["columns"]?.GetValue<int>() ?? throw new ArgumentException("columns is required");
        var dataArray = arguments?["data"]?.AsArray();

        using var document = new Document(path);
        var page = document.Pages[pageIndex];

        var table = new Table
        {
            ColumnWidths = string.Join(" ", Enumerable.Repeat("100", columns)),
            Border = new BorderInfo(BorderSide.All, 1f),
            DefaultCellBorder = new BorderInfo(BorderSide.All, 0.5f)
        };

        for (int i = 0; i < rows; i++)
        {
            var row = table.Rows.Add();
            for (int j = 0; j < columns; j++)
            {
                var cell = row.Cells.Add();
                
                if (dataArray != null && i < dataArray.Count)
                {
                    var rowArray = dataArray[i]?.AsArray();
                    if (rowArray != null && j < rowArray.Count)
                    {
                        cell.Paragraphs.Add(new Aspose.Pdf.Text.TextFragment(rowArray[j]?.GetValue<string>() ?? ""));
                    }
                }
            }
        }

        page.Paragraphs.Add(table);
        document.Save(path);

        return await Task.FromResult($"Table ({rows}x{columns}) added to page {pageIndex}: {path}");
    }
}

