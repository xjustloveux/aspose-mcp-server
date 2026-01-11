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
        var tableIndex = parameters.GetOptional("tableIndex", 0);
        var cellRow = parameters.GetOptional<int?>("cellRow");
        var cellColumn = parameters.GetOptional<int?>("cellColumn");
        var cellValue = parameters.GetOptional<string?>("cellValue");

        var document = context.Document;

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

        if (tables.Count == 0)
            try
            {
                var tablesFromLinq = document.Pages
                    .SelectMany(p => p.Paragraphs?.OfType<Aspose.Pdf.Table>() ?? Enumerable.Empty<Aspose.Pdf.Table>())
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
                                if (paragraph is Aspose.Pdf.Table foundTable) tables.Add(foundTable);
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

        MarkModified(context);

        return Success($"Edited table {tableIndex}.");
    }
}
