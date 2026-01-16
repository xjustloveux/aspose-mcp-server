using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Table;

/// <summary>
///     Handler for editing tables in PDF documents.
/// </summary>
public class EditPdfTableHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits an existing table in the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: tableIndex
    ///     Optional: cellRow, cellColumn, cellValue
    /// </param>
    /// <returns>Success message with table edit details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractEditParameters(parameters);

        var document = context.Document;
        var tables = FindTables(document);

        if (tables.Count == 0)
            throw new ArgumentException(BuildNoTablesErrorMessage(document));

        if (p.TableIndex < 0 || p.TableIndex >= tables.Count)
            throw new ArgumentException(
                $"tableIndex must be between 0 and {tables.Count - 1} (found {tables.Count} table(s))");

        var table = tables[p.TableIndex];
        EditTableCell(table, p.CellRow, p.CellColumn, p.CellValue);

        MarkModified(context);

        return Success($"Edited table {p.TableIndex}.");
    }

    /// <summary>
    ///     Extracts parameters for edit operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static EditParameters ExtractEditParameters(OperationParameters parameters)
    {
        return new EditParameters(
            parameters.GetOptional("tableIndex", 0),
            parameters.GetOptional<int?>("cellRow"),
            parameters.GetOptional<int?>("cellColumn"),
            parameters.GetOptional<string?>("cellValue")
        );
    }

    /// <summary>
    ///     Finds all tables in the PDF document using multiple search methods.
    /// </summary>
    /// <param name="document">The PDF document to search.</param>
    /// <returns>A list of tables found in the document.</returns>
    private static List<Aspose.Pdf.Table> FindTables(Document document)
    {
        var tables = FindTablesUsingForeach(document);
        if (tables.Count == 0)
            tables = FindTablesUsingLinq(document);
        if (tables.Count == 0)
            tables = FindTablesUsingIndexer(document);
        return tables;
    }

    /// <summary>
    ///     Finds tables using foreach iteration over paragraphs.
    /// </summary>
    /// <param name="document">The PDF document to search.</param>
    /// <returns>A list of tables found using the foreach method.</returns>
    private static List<Aspose.Pdf.Table> FindTablesUsingForeach(Document document)
    {
        List<Aspose.Pdf.Table> tables = [];
        foreach (var page in document.Pages)
            try
            {
                var paragraphs = page.Paragraphs;
                if (paragraphs is { Count: > 0 })
                    foreach (var paragraph in paragraphs)
                        if (paragraph is Aspose.Pdf.Table foundTable)
                            tables.Add(foundTable);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[WARN] Error searching tables in document: {ex.Message}");
            }

        return tables;
    }

    /// <summary>
    ///     Finds tables using LINQ queries.
    /// </summary>
    /// <param name="document">The PDF document to search.</param>
    /// <returns>A list of tables found using LINQ.</returns>
    private static List<Aspose.Pdf.Table> FindTablesUsingLinq(Document document)
    {
        try
        {
            return document.Pages
                .SelectMany(p => p.Paragraphs?.OfType<Aspose.Pdf.Table>() ?? Enumerable.Empty<Aspose.Pdf.Table>())
                .ToList();
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[WARN] Error using LINQ method to find tables: {ex.Message}");
            return [];
        }
    }

    /// <summary>
    ///     Finds tables using index-based paragraph access.
    /// </summary>
    /// <param name="document">The PDF document to search.</param>
    /// <returns>A list of tables found using indexer access.</returns>
    private static List<Aspose.Pdf.Table> FindTablesUsingIndexer(Document document)
    {
        List<Aspose.Pdf.Table> tables = [];
        try
        {
            foreach (var page in document.Pages)
            {
                var paragraphs = page.Paragraphs;
                if (paragraphs is not { Count: > 0 }) continue;

                for (var i = 1; i <= paragraphs.Count; i++)
                    try
                    {
                        if (paragraphs[i] is Aspose.Pdf.Table foundTable)
                            tables.Add(foundTable);
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine($"[WARN] Error accessing paragraph at index {i}: {ex.Message}");
                    }
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[WARN] Error in alternative table search method: {ex.Message}");
        }

        return tables;
    }

    /// <summary>
    ///     Builds a detailed error message when no tables are found in the document.
    /// </summary>
    /// <param name="document">The PDF document that was searched.</param>
    /// <returns>A detailed error message with diagnostic information.</returns>
    private static string BuildNoTablesErrorMessage(Document document)
    {
        var (totalParagraphs, paragraphTypes, pageInfo) = AnalyzeDocumentStructure(document);

        var typeInfo = paragraphTypes.Count > 0
            ? $" Paragraph types found: {string.Join(", ", paragraphTypes.Select(kvp => $"{kvp.Key}({kvp.Value})"))}"
            : string.Empty;

        var pageDetails = pageInfo.Count > 0 ? $" Page details: {string.Join("; ", pageInfo)}" : "";

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

        return errorMsg;
    }

    /// <summary>
    ///     Analyzes the structure of the PDF document for diagnostic purposes.
    /// </summary>
    /// <param name="document">The PDF document to analyze.</param>
    /// <returns>A tuple containing paragraph count, paragraph types, and page information.</returns>
    private static (int totalParagraphs, Dictionary<string, int> paragraphTypes, List<string> pageInfo)
        AnalyzeDocumentStructure(Document document)
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

        return (totalParagraphs, paragraphTypes, pageInfo);
    }

    /// <summary>
    ///     Edits the content of a specific cell in a table.
    /// </summary>
    /// <param name="table">The table to edit.</param>
    /// <param name="cellRow">The 0-based row index of the cell.</param>
    /// <param name="cellColumn">The 0-based column index of the cell.</param>
    /// <param name="cellValue">The new value to set for the cell.</param>
    private static void EditTableCell(Aspose.Pdf.Table table, int? cellRow, int? cellColumn, string? cellValue)
    {
        if (!cellRow.HasValue || !cellColumn.HasValue || string.IsNullOrEmpty(cellValue))
            return;

        if (cellRow.Value < 0 || cellRow.Value >= table.Rows.Count)
            throw new ArgumentException($"cellRow must be between 0 and {table.Rows.Count - 1}");
        if (cellColumn.Value < 0 || cellColumn.Value >= table.Rows[cellRow.Value].Cells.Count)
            throw new ArgumentException(
                $"cellColumn must be between 0 and {table.Rows[cellRow.Value].Cells.Count - 1}");

        var cell = table.Rows[cellRow.Value].Cells[cellColumn.Value];
        cell.Paragraphs.Clear();
        cell.Paragraphs.Add(new TextFragment(cellValue));
    }

    /// <summary>
    ///     Parameters for edit operation.
    /// </summary>
    /// <param name="TableIndex">The 0-based table index.</param>
    /// <param name="CellRow">The optional cell row index.</param>
    /// <param name="CellColumn">The optional cell column index.</param>
    /// <param name="CellValue">The optional cell value.</param>
    private record EditParameters(int TableIndex, int? CellRow, int? CellColumn, string? CellValue);
}
