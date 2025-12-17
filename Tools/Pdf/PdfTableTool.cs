using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Unified tool for managing PDF tables (add, edit)
/// </summary>
public class PdfTableTool : IAsposeTool
{
    public string Description => @"Manage tables in PDF documents. Supports 2 operations: add, edit.

Usage examples:
- Add table: pdf_table(operation='add', path='doc.pdf', pageIndex=1, rows=3, columns=3, data=[['A','B','C'],['1','2','3']], x=100, y=100)
- Edit table: pdf_table(operation='edit', path='doc.pdf', pageIndex=1, tableIndex=0, data=[['X','Y','Z']])";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add a table (required params: path, pageIndex, rows, columns, data)
- 'edit': Edit table data (required params: path, pageIndex, tableIndex, data)",
                @enum = new[] { "add", "edit" }
            },
            path = new
            {
                type = "string",
                description = "PDF file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to overwrite input)"
            },
            pageIndex = new
            {
                type = "number",
                description = "Page index (1-based, required for add)"
            },
            rows = new
            {
                type = "number",
                description = "Number of rows (required for add)"
            },
            columns = new
            {
                type = "number",
                description = "Number of columns (required for add)"
            },
            data = new
            {
                type = "array",
                description = "Table data (array of arrays, for add)",
                items = new
                {
                    type = "array",
                    items = new { type = "string" }
                }
            },
            x = new
            {
                type = "number",
                description = "X position (for add, default: 100)"
            },
            y = new
            {
                type = "number",
                description = "Y position (for add, default: 600)"
            },
            tableIndex = new
            {
                type = "number",
                description = "Table index (0-based, required for edit)"
            },
            cellRow = new
            {
                type = "number",
                description = "Cell row index (0-based, for edit)"
            },
            cellColumn = new
            {
                type = "number",
                description = "Cell column index (0-based, for edit)"
            },
            cellValue = new
            {
                type = "string",
                description = "New cell value (for edit)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");

        return operation.ToLower() switch
        {
            "add" => await AddTable(arguments),
            "edit" => await EditTable(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a table to a PDF page
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, pageIndex, rows, columns, data, x, y, optional outputPath</param>
    /// <returns>Success message</returns>
    private async Task<string> AddTable(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");
        var rows = ArgumentHelper.GetInt(arguments, "rows");
        var columns = ArgumentHelper.GetInt(arguments, "columns");
        _ = ArgumentHelper.GetDouble(arguments, "x", "x", false, 100);
        _ = ArgumentHelper.GetDouble(arguments, "y", "y", false, 600);

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        var table = new Table
        {
            ColumnWidths = string.Join(" ", Enumerable.Repeat("100", columns)),
            DefaultCellBorder = new BorderInfo(BorderSide.All, 0.5F)
        };

        string[][]? data = null;
        if (arguments?.ContainsKey("data") == true)
            try
            {
                var dataJson = arguments["data"]?.ToJsonString();
                if (!string.IsNullOrEmpty(dataJson))
                    data = JsonSerializer.Deserialize<string[][]>(dataJson);
            }
            catch (Exception jsonEx)
            {
                throw new ArgumentException(
                    $"Unable to parse data parameter: {jsonEx.Message}. Please ensure data is a valid 2D string array format, e.g.: [[\"A1\",\"B1\"],[\"A2\",\"B2\"]]");
            }

        for (var i = 0; i < rows; i++)
        {
            var row = table.Rows.Add();
            for (var j = 0; j < columns; j++)
            {
                var cell = row.Cells.Add();
                var cellText = data != null && i < data.Length && j < data[i].Length
                    ? data[i][j]
                    : $"Cell {i + 1},{j + 1}";
                cell.Paragraphs.Add(new TextFragment(cellText));
            }
        }

        page.Paragraphs.Add(table);
        document.Save(outputPath);
        return await Task.FromResult(
            $"Successfully added table ({rows} rows x {columns} columns) to page {pageIndex}. Output: {outputPath}");
    }

    /// <summary>
    ///     Edits table data in a PDF
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, pageIndex, tableIndex, data, optional outputPath</param>
    /// <returns>Success message</returns>
    private async Task<string> EditTable(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var tableIndex = ArgumentHelper.GetInt(arguments, "tableIndex");
        var cellRow = ArgumentHelper.GetIntNullable(arguments, "cellRow");
        var cellColumn = ArgumentHelper.GetIntNullable(arguments, "cellColumn");
        var cellValue = ArgumentHelper.GetStringNullable(arguments, "cellValue");

        SecurityHelper.ValidateFilePath(path);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        using var document = new Document(path);
        var tables = document.Pages
            .SelectMany(p => p.Paragraphs.OfType<Table>())
            .ToList();

        if (tableIndex < 0 || tableIndex >= tables.Count)
            throw new ArgumentException($"tableIndex must be between 0 and {tables.Count - 1}");

        var table = tables[tableIndex];

        if (cellRow.HasValue && cellColumn.HasValue && !string.IsNullOrEmpty(cellValue))
        {
            if (cellRow.Value < 0 || cellRow.Value >= table.Rows.Count)
                throw new ArgumentException($"cellRow must be between 0 and {table.Rows.Count - 1}");
            if (cellColumn.Value < 0 || cellColumn.Value >= table.Rows[cellRow.Value].Cells.Count)
                throw new ArgumentException(
                    $"cellColumn must be between 0 and {table.Rows[cellRow.Value].Cells.Count - 1}");

            var cell = table.Rows[cellRow.Value].Cells[cellColumn.Value];
            cell.Paragraphs.Clear();
            cell.Paragraphs.Add(new TextFragment(cellValue));
        }

        document.Save(outputPath);
        return await Task.FromResult($"Successfully edited table {tableIndex}. Output: {outputPath}");
    }
}