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
- Add table: pdf_table(operation='add', path='doc.pdf', pageIndex=1, rows=3, columns=3, data=[['A','B','C'],['1','2','3']])
- Add table with position: pdf_table(operation='add', path='doc.pdf', pageIndex=1, rows=2, columns=2, x=100, y=500)
- Add table with column widths: pdf_table(operation='add', path='doc.pdf', pageIndex=1, rows=2, columns=3, columnWidths='100 150 200')
- Edit table cell: pdf_table(operation='edit', path='doc.pdf', tableIndex=0, cellRow=0, cellColumn=1, cellValue='NewValue')

Note: PDF table editing has limitations. After saving, tables may be converted to graphics and cannot be edited as Table objects.";

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
                description = "X position (left margin) in PDF points (for add, default: 100)"
            },
            y = new
            {
                type = "number",
                description = "Y position (top margin) in PDF points (for add, default: 600)"
            },
            columnWidths = new
            {
                type = "string",
                description =
                    "Space-separated column widths in PDF points (for add, default: all columns 100pt, e.g., '100 150 200')"
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
    /// <exception cref="ArgumentException">Thrown when operation is unknown or required parameters are missing</exception>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        return operation.ToLower() switch
        {
            "add" => await AddTable(path, outputPath, arguments),
            "edit" => await EditTable(path, outputPath, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a table to a PDF page
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing pageIndex, rows, columns, optional data, x, y, columnWidths</param>
    /// <returns>Success message</returns>
    /// <exception cref="ArgumentException">Thrown when pageIndex is out of range or data format is invalid</exception>
    private Task<string> AddTable(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");
            var rows = ArgumentHelper.GetInt(arguments, "rows");
            var columns = ArgumentHelper.GetInt(arguments, "columns");
            var x = ArgumentHelper.GetDouble(arguments, "x", "x", false, 100);
            var y = ArgumentHelper.GetDouble(arguments, "y", "y", false, 600);
            var columnWidths = ArgumentHelper.GetStringNullable(arguments, "columnWidths");

            using var document = new Document(path);
            if (pageIndex < 1 || pageIndex > document.Pages.Count)
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

            var page = document.Pages[pageIndex];

            // Determine column widths: use provided value or default to 100 for each column
            var effectiveColumnWidths = !string.IsNullOrEmpty(columnWidths)
                ? columnWidths
                : string.Join(" ", Enumerable.Repeat("100", columns));

            var table = new Table
            {
                ColumnWidths = effectiveColumnWidths,
                DefaultCellBorder = new BorderInfo(BorderSide.All, 0.5F),
                Margin = new MarginInfo { Left = x, Top = y }
            };

            // Parse data using direct deserialization for better performance
            string[][]? data = null;
            if (arguments?.ContainsKey("data") == true && arguments["data"] != null)
                try
                {
                    data = arguments["data"].Deserialize<string[][]>();
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
            return $"Added table ({rows} rows x {columns} columns) to page {pageIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits table data in a PDF
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing tableIndex, optional cellRow, cellColumn, cellValue</param>
    /// <returns>Success message</returns>
    /// <exception cref="ArgumentException">Thrown when no tables found, tableIndex out of range, or cell indices out of range</exception>
    private Task<string> EditTable(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var tableIndex = ArgumentHelper.GetInt(arguments, "tableIndex");
            var cellRow = ArgumentHelper.GetIntNullable(arguments, "cellRow");
            var cellColumn = ArgumentHelper.GetIntNullable(arguments, "cellColumn");
            var cellValue = ArgumentHelper.GetStringNullable(arguments, "cellValue");

            using var document = new Document(path);

            var tables = new List<Table>();

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
                    // Continue searching
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
                    // LINQ method failed, continue with empty list
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
                                    // Skip invalid indices
                                    Console.Error.WriteLine(
                                        $"[WARN] Error accessing paragraph at index {i}: {ex.Message}");
                                }
                    }
                }
                catch (Exception ex)
                {
                    // Method 3 failed
                    Console.Error.WriteLine($"[WARN] Error in alternative table search method: {ex.Message}");
                }

            if (tables.Count == 0)
            {
                // Provide more detailed error message with debugging information
                var totalParagraphs = 0;
                var paragraphTypes = new Dictionary<string, int>();
                var pageInfo = new List<string>();

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
                    // Error counting paragraphs
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

                // Additional note about Aspose.Pdf limitation
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

            document.Save(outputPath);
            return $"Edited table {tableIndex}. Output: {outputPath}";
        });
    }
}