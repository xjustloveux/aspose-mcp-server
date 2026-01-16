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
        var p = ExtractAddParameters(parameters);

        if (p.Rows <= 0)
            throw new ArgumentException("rows is required and must be greater than 0 for add operation");
        if (p.Columns <= 0)
            throw new ArgumentException("columns is required and must be greater than 0 for add operation");

        var document = context.Document;

        if (p.PageIndex < 1 || p.PageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[p.PageIndex];

        var effectiveColumnWidths = !string.IsNullOrEmpty(p.ColumnWidths)
            ? p.ColumnWidths
            : string.Join(" ", Enumerable.Repeat("100", p.Columns));

        var table = new Aspose.Pdf.Table
        {
            ColumnWidths = effectiveColumnWidths,
            DefaultCellBorder = new BorderInfo(BorderSide.All, 0.5F),
            Margin = new MarginInfo { Left = p.X, Top = p.Y }
        };

        for (var i = 0; i < p.Rows; i++)
        {
            var row = table.Rows.Add();
            for (var j = 0; j < p.Columns; j++)
            {
                var cell = row.Cells.Add();
                var cellText = p.Data != null && i < p.Data.Length && j < p.Data[i].Length
                    ? p.Data[i][j]
                    : $"Cell {i + 1},{j + 1}";
                cell.Paragraphs.Add(new TextFragment(cellText));
            }
        }

        page.Paragraphs.Add(table);

        MarkModified(context);

        return Success($"Added table ({p.Rows} rows x {p.Columns} columns) to page {p.PageIndex}.");
    }

    /// <summary>
    ///     Extracts parameters for add operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static AddParameters ExtractAddParameters(OperationParameters parameters)
    {
        return new AddParameters(
            parameters.GetOptional("pageIndex", 1),
            parameters.GetOptional("rows", 0),
            parameters.GetOptional("columns", 0),
            parameters.GetOptional<string[][]?>("data"),
            parameters.GetOptional("x", 100.0),
            parameters.GetOptional("y", 600.0),
            parameters.GetOptional<string?>("columnWidths")
        );
    }

    /// <summary>
    ///     Parameters for add operation.
    /// </summary>
    /// <param name="PageIndex">The 1-based page index.</param>
    /// <param name="Rows">The number of rows.</param>
    /// <param name="Columns">The number of columns.</param>
    /// <param name="Data">The optional table data.</param>
    /// <param name="X">The X position.</param>
    /// <param name="Y">The Y position.</param>
    /// <param name="ColumnWidths">The optional column widths.</param>
    private record AddParameters(
        int PageIndex,
        int Rows,
        int Columns,
        string[][]? Data,
        double X,
        double Y,
        string? ColumnWidths);
}
