using Aspose.Words;
using Aspose.Words.Tables;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Table;

/// <summary>
///     Handler for splitting cells in Word document tables.
/// </summary>
public class SplitCellWordTableHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "split_cell";

    /// <summary>
    ///     Splits a cell in a table.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: rowIndex, columnIndex.
    ///     Optional: tableIndex (default 0), splitRows (default 2), splitCols (default 2), sectionIndex.
    /// </param>
    /// <returns>Success message with split dimensions.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or indices are out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when trying to split a merged cell.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var rowIndex = parameters.GetOptional<int?>("rowIndex");
        var columnIndex = parameters.GetOptional<int?>("columnIndex");

        if (!rowIndex.HasValue)
            throw new ArgumentException("rowIndex is required for split_cell operation");
        if (!columnIndex.HasValue)
            throw new ArgumentException("columnIndex is required for split_cell operation");

        var tableIndex = parameters.GetOptional("tableIndex", 0);
        var splitRows = parameters.GetOptional("splitRows", 2);
        var splitCols = parameters.GetOptional("splitCols", 2);
        var sectionIndex = parameters.GetOptional<int?>("sectionIndex");

        var doc = context.Document;
        var actualSectionIndex = sectionIndex ?? 0;
        if (actualSectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"Section index {actualSectionIndex} out of range");

        var section = doc.Sections[actualSectionIndex];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Aspose.Words.Tables.Table>().ToList();
        if (tableIndex < 0 || tableIndex >= tables.Count)
            throw new ArgumentException($"Table index {tableIndex} out of range");

        var table = tables[tableIndex];
        if (rowIndex.Value < 0 || rowIndex.Value >= table.Rows.Count)
            throw new ArgumentException($"Row index {rowIndex.Value} out of range");

        var row = table.Rows[rowIndex.Value];
        if (columnIndex.Value < 0 || columnIndex.Value >= row.Cells.Count)
            throw new ArgumentException($"Column index {columnIndex.Value} out of range");

        var cell = row.Cells[columnIndex.Value];
        var isMerged = cell.CellFormat.HorizontalMerge != CellMerge.None ||
                       cell.CellFormat.VerticalMerge != CellMerge.None;
        if (isMerged)
            throw new InvalidOperationException("Cannot split merged cell. Please unmerge first or edit directly.");

        var cellText = cell.GetText();
        var parentRow = cell.ParentRow;
        var cellIndex = parentRow.Cells.IndexOf(cell);

        SplitCellHorizontally(doc, parentRow, cell, cellIndex, cellText, splitCols);

        if (splitRows > 1)
            AddSplitRows(doc, table, rowIndex.Value, splitRows);

        MarkModified(context);

        return Success(
            $"Successfully split cell [{rowIndex.Value}, {columnIndex.Value}] into {splitRows} rows x {splitCols} columns.");
    }

    /// <summary>
    ///     Splits a cell horizontally into multiple columns.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="parentRow">The parent row.</param>
    /// <param name="cell">The cell to split.</param>
    /// <param name="cellIndex">The cell index.</param>
    /// <param name="cellText">The original cell text.</param>
    /// <param name="splitCols">Number of columns to split into.</param>
    private static void SplitCellHorizontally(Document doc, Row parentRow, Cell cell,
        int cellIndex, string cellText, int splitCols)
    {
        for (var c = 0; c < splitCols; c++)
        {
            var newCell = new Cell(doc)
            {
                CellFormat =
                {
                    Width = cell.CellFormat.Width / splitCols,
                    VerticalAlignment = cell.CellFormat.VerticalAlignment
                }
            };
            newCell.CellFormat.SetPaddings(
                cell.CellFormat.TopPadding,
                cell.CellFormat.BottomPadding,
                cell.CellFormat.LeftPadding,
                cell.CellFormat.RightPadding
            );

            var para = new WordParagraph(doc);
            if (splitCols == 1 || (c == 0 && !string.IsNullOrEmpty(cellText)))
            {
                var run = new Run(doc, cellText);
                para.AppendChild(run);
            }

            newCell.AppendChild(para);

            if (c == 0)
            {
                parentRow.Cells[cellIndex].Remove();
                parentRow.Cells.Insert(cellIndex, newCell);
            }
            else
            {
                parentRow.Cells.Insert(cellIndex + c, newCell);
            }
        }
    }

    /// <summary>
    ///     Adds additional rows for vertical split.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="table">The table.</param>
    /// <param name="rowIndex">The original row index.</param>
    /// <param name="splitRows">Number of rows to split into.</param>
    private static void AddSplitRows(Document doc, Aspose.Words.Tables.Table table, int rowIndex, int splitRows)
    {
        for (var r = 1; r < splitRows; r++)
        {
            var insertAfterRowIndex = rowIndex + r - 1;
            if (insertAfterRowIndex < table.Rows.Count)
            {
                var newRow = new Row(doc);
                var sourceRow = table.Rows[rowIndex];
                for (var c = 0; c < sourceRow.Cells.Count; c++)
                {
                    var newCellInRow = new Cell(doc);
                    var sourceCell = sourceRow.Cells[c];
                    newCellInRow.CellFormat.Width = sourceCell.CellFormat.Width;
                    newCellInRow.CellFormat.VerticalAlignment = sourceCell.CellFormat.VerticalAlignment;
                    newCellInRow.CellFormat.SetPaddings(
                        sourceCell.CellFormat.TopPadding,
                        sourceCell.CellFormat.BottomPadding,
                        sourceCell.CellFormat.LeftPadding,
                        sourceCell.CellFormat.RightPadding
                    );
                    var paraInRow = new WordParagraph(doc);
                    newCellInRow.AppendChild(paraInRow);
                    newRow.AppendChild(newCellInRow);
                }

                table.InsertAfter(newRow, table.Rows[insertAfterRowIndex]);
            }
        }
    }
}
