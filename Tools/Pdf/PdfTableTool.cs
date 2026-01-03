using System.ComponentModel;
using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Unified tool for managing PDF tables (add, edit)
/// </summary>
[McpServerToolType]
public class PdfTableTool
{
    /// <summary>
    ///     The document session manager for managing in-memory document sessions.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PdfTableTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    public PdfTableTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "pdf_table")]
    [Description(@"Manage tables in PDF documents. Supports 2 operations: add, edit.

Usage examples:
- Add table: pdf_table(operation='add', path='doc.pdf', pageIndex=1, rows=3, columns=3, data=[['A','B','C'],['1','2','3']])
- Add table with position: pdf_table(operation='add', path='doc.pdf', pageIndex=1, rows=2, columns=2, x=100, y=500)
- Add table with column widths: pdf_table(operation='add', path='doc.pdf', pageIndex=1, rows=2, columns=3, columnWidths='100 150 200')
- Edit table cell: pdf_table(operation='edit', path='doc.pdf', tableIndex=0, cellRow=0, cellColumn=1, cellValue='NewValue')

Note: PDF table editing has limitations. After saving, tables may be converted to graphics and cannot be edited as Table objects.")]
    public string Execute(
        [Description("Operation: add, edit")] string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Page index (1-based, required for add)")]
        int pageIndex = 1,
        [Description("Number of rows (required for add)")]
        int rows = 0,
        [Description("Number of columns (required for add)")]
        int columns = 0,
        [Description("Table data (array of arrays, for add)")]
        string[][]? data = null,
        [Description("X position (left margin) in PDF points (for add, default: 100)")]
        double x = 100,
        [Description("Y position (top margin) in PDF points (for add, default: 600)")]
        double y = 600,
        [Description("Space-separated column widths in PDF points (for add, e.g., '100 150 200')")]
        string? columnWidths = null,
        [Description("Table index (0-based, required for edit)")]
        int tableIndex = 0,
        [Description("Cell row index (0-based, for edit)")]
        int? cellRow = null,
        [Description("Cell column index (0-based, for edit)")]
        int? cellColumn = null,
        [Description("New cell value (for edit)")]
        string? cellValue = null)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);

        return operation.ToLower() switch
        {
            "add" => AddTable(ctx, outputPath, pageIndex, rows, columns, data, x, y, columnWidths),
            "edit" => EditTable(ctx, outputPath, tableIndex, cellRow, cellColumn, cellValue),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a new table to the specified page of the PDF document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="pageIndex">The 1-based page index.</param>
    /// <param name="rows">Number of rows in the table.</param>
    /// <param name="columns">Number of columns in the table.</param>
    /// <param name="data">Optional table data as an array of arrays.</param>
    /// <param name="x">The X position (left margin) in PDF points.</param>
    /// <param name="y">The Y position (top margin) in PDF points.</param>
    /// <param name="columnWidths">Optional space-separated column widths.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    private static string AddTable(DocumentContext<Document> ctx, string? outputPath,
        int pageIndex, int rows, int columns, string[][]? data, double x, double y, string? columnWidths)
    {
        if (rows <= 0)
            throw new ArgumentException("rows is required and must be greater than 0 for add operation");
        if (columns <= 0)
            throw new ArgumentException("columns is required and must be greater than 0 for add operation");

        var document = ctx.Document;

        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];

        var effectiveColumnWidths = !string.IsNullOrEmpty(columnWidths)
            ? columnWidths
            : string.Join(" ", Enumerable.Repeat("100", columns));

        var table = new Table
        {
            ColumnWidths = effectiveColumnWidths,
            DefaultCellBorder = new BorderInfo(BorderSide.All, 0.5F),
            Margin = new MarginInfo { Left = x, Top = y }
        };

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

        ctx.Save(outputPath);

        return $"Added table ({rows} rows x {columns} columns) to page {pageIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Edits an existing table in the PDF document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="tableIndex">The 0-based table index.</param>
    /// <param name="cellRow">Optional 0-based row index for cell editing.</param>
    /// <param name="cellColumn">Optional 0-based column index for cell editing.</param>
    /// <param name="cellValue">Optional new value for the specified cell.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when no tables are found or indices are invalid.</exception>
    private static string EditTable(DocumentContext<Document> ctx, string? outputPath,
        int tableIndex, int? cellRow, int? cellColumn, string? cellValue)
    {
        var document = ctx.Document;

        List<Table> tables = [];

        foreach (var page in document.Pages)
            try
            {
                var paragraphs = page.Paragraphs;
                if (paragraphs is { Count: > 0 })
                    foreach (var paragraph in paragraphs)
                        if (paragraph is Table foundTable)
                            tables.Add(foundTable);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[WARN] Error searching tables in document: {ex.Message}");
            }

        if (tables.Count == 0)
            try
            {
                var tablesFromLinq = document.Pages
                    .SelectMany(p => p.Paragraphs?.OfType<Table>() ?? Enumerable.Empty<Table>())
                    .ToList();
                if (tablesFromLinq.Count > 0) tables.AddRange(tablesFromLinq);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[WARN] Error using LINQ method to find tables: {ex.Message}");
            }

        if (tables.Count == 0)
            try
            {
                foreach (var page in document.Pages)
                {
                    var paragraphs = page.Paragraphs;
                    if (paragraphs is { Count: > 0 })
                        for (var i = 1; i <= paragraphs.Count; i++)
                            try
                            {
                                var paragraph = paragraphs[i];
                                if (paragraph is Table foundTable) tables.Add(foundTable);
                            }
                            catch (Exception ex)
                            {
                                Console.Error.WriteLine(
                                    $"[WARN] Error accessing paragraph at index {i}: {ex.Message}");
                            }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[WARN] Error in alternative table search method: {ex.Message}");
            }

        if (tables.Count == 0)
        {
            var totalParagraphs = 0;
            var paragraphTypes = new Dictionary<string, int>();
            List<string> pageInfo = [];

            try
            {
                for (var pageNum = 1; pageNum <= document.Pages.Count; pageNum++)
                {
                    var page = document.Pages[pageNum];
                    var pageParagraphCount = page.Paragraphs?.Count ?? 0;
                    totalParagraphs += pageParagraphCount;

                    if (pageParagraphCount > 0 && page.Paragraphs != null)
                    {
                        pageInfo.Add($"Page {pageNum}: {pageParagraphCount} paragraphs");

                        foreach (var paragraph in page.Paragraphs)
                        {
                            var typeName = paragraph.GetType().Name;
                            paragraphTypes[typeName] = paragraphTypes.GetValueOrDefault(typeName, 0) + 1;
                        }
                    }
                    else
                    {
                        pageInfo.Add($"Page {pageNum}: 0 paragraphs");
                    }
                }
            }
            catch (Exception ex)
            {
                pageInfo.Add($"Error analyzing pages: {ex.Message}");
            }

            var typeInfo = paragraphTypes.Count > 0
                ? $" Paragraph types found: {string.Join(", ", paragraphTypes.Select(kvp => $"{kvp.Key}({kvp.Value})"))}"
                : string.Empty;

            var pageDetails = pageInfo.Count > 0
                ? $" Page details: {string.Join("; ", pageInfo)}"
                : "";

            var errorMsg =
                $"No tables found in the document. Total paragraphs across all pages: {totalParagraphs}.{typeInfo}{pageDetails}";
            errorMsg +=
                " Make sure tables are added using the 'add' operation first, and that the document has been saved after adding tables.";
            errorMsg +=
                " Note: If you just added a table, ensure you're editing the saved output file, not the original input file.";

            if (totalParagraphs == 0)
            {
                errorMsg +=
                    " IMPORTANT: After saving, tables may be converted to graphics objects and cannot be edited as Table objects.";
                errorMsg += " This is a limitation of the PDF format and Aspose.Pdf library.";
                errorMsg += " To edit tables, you may need to recreate them or use a different approach.";
            }

            throw new ArgumentException(errorMsg);
        }

        if (tableIndex < 0 || tableIndex >= tables.Count)
            throw new ArgumentException(
                $"tableIndex must be between 0 and {tables.Count - 1} (found {tables.Count} table(s))");

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

        ctx.Save(outputPath);

        return $"Edited table {tableIndex}. {ctx.GetOutputMessage(outputPath)}";
    }
}