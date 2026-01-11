using Aspose.Pdf;
using Aspose.Pdf.Text;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Table;

/// <summary>
///     Handler for adding tables to PDF documents.
/// </summary>
public class AddPdfTableHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a new table to the specified page of the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: pageIndex, rows, columns
    ///     Optional: data, x, y, columnWidths
    /// </param>
    /// <returns>Success message with table creation details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var pageIndex = parameters.GetOptional("pageIndex", 1);
        var rows = parameters.GetOptional("rows", 0);
        var columns = parameters.GetOptional("columns", 0);
        var data = parameters.GetOptional<string[][]?>("data");
        var x = parameters.GetOptional("x", 100.0);
        var y = parameters.GetOptional("y", 600.0);
        var columnWidths = parameters.GetOptional<string?>("columnWidths");

        if (rows <= 0)
            throw new ArgumentException("rows is required and must be greater than 0 for add operation");
        if (columns <= 0)
            throw new ArgumentException("columns is required and must be greater than 0 for add operation");

        var document = context.Document;

        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];

        var effectiveColumnWidths = !string.IsNullOrEmpty(columnWidths)
            ? columnWidths
            : string.Join(" ", Enumerable.Repeat("100", columns));

        var table = new Aspose.Pdf.Table
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

        MarkModified(context);

        return Success($"Added table ({rows} rows x {columns} columns) to page {pageIndex}.");
    }
}
