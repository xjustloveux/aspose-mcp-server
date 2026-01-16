using Aspose.Words;
using Aspose.Words.Tables;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Table;

/// <summary>
///     Handler for merging cells in Word document tables.
/// </summary>
public class MergeCellsWordTableHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "merge_cells";

    /// <summary>
    ///     Merges cells in a table.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: startRow, startCol, endRow, endCol.
    ///     Optional: tableIndex (default 0), sectionIndex.
    /// </param>
    /// <returns>Success message with merged cell range.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or indices are out of range.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractMergeCellsParameters(parameters);

        var doc = context.Document;
        var actualSectionIndex = p.SectionIndex ?? 0;
        if (actualSectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"Section index {actualSectionIndex} out of range");

        var section = doc.Sections[actualSectionIndex];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Aspose.Words.Tables.Table>().ToList();
        if (p.TableIndex < 0 || p.TableIndex >= tables.Count)
            throw new ArgumentException($"Table index {p.TableIndex} out of range");

        var table = tables[p.TableIndex];

        ValidateMergeRange(table, p.StartRow, p.EndRow, p.StartCol, p.EndCol);

        ApplyMerge(table, p.StartRow, p.EndRow, p.StartCol, p.EndCol);

        MarkModified(context);

        return Success(
            $"Successfully merged cells from [{p.StartRow}, {p.StartCol}] to [{p.EndRow}, {p.EndCol}].");
    }

    /// <summary>
    ///     Validates the merge range is within table bounds.
    /// </summary>
    /// <param name="table">The table.</param>
    /// <param name="startRow">Start row index.</param>
    /// <param name="endRow">End row index.</param>
    /// <param name="startCol">Start column index.</param>
    /// <param name="endCol">End column index.</param>
    /// <exception cref="ArgumentException">Thrown when indices are out of range.</exception>
    private static void ValidateMergeRange(Aspose.Words.Tables.Table table, int startRow, int endRow, int startCol,
        int endCol)
    {
        if (startRow < 0 || startRow >= table.Rows.Count || endRow < 0 || endRow >= table.Rows.Count)
            throw new ArgumentException("Row indices out of range");

        if (startRow > endRow)
            throw new ArgumentException($"Start row {startRow} cannot be greater than end row {endRow}");

        var firstRowForCheck = table.Rows[startRow];
        if (startCol < 0 || startCol >= firstRowForCheck.Cells.Count || endCol < 0 ||
            endCol >= firstRowForCheck.Cells.Count)
            throw new ArgumentException("Column indices out of range");

        if (startCol > endCol)
            throw new ArgumentException($"Start column {startCol} cannot be greater than end column {endCol}");
    }

    /// <summary>
    ///     Applies merge to the specified cell range.
    /// </summary>
    /// <param name="table">The table.</param>
    /// <param name="startRow">Start row index.</param>
    /// <param name="endRow">End row index.</param>
    /// <param name="startCol">Start column index.</param>
    /// <param name="endCol">End column index.</param>
    private static void ApplyMerge(Aspose.Words.Tables.Table table, int startRow, int endRow, int startCol, int endCol)
    {
        for (var row = startRow; row <= endRow; row++)
        {
            var currentRow = table.Rows[row];
            for (var col = startCol; col <= endCol; col++)
            {
                var cell = currentRow.Cells[col];
                if (row == startRow && col == startCol)
                {
                    if (startRow != endRow)
                        cell.CellFormat.VerticalMerge = CellMerge.First;
                    if (startCol != endCol)
                        cell.CellFormat.HorizontalMerge = CellMerge.First;
                }
                else
                {
                    if (row == startRow)
                    {
                        cell.CellFormat.HorizontalMerge = CellMerge.Previous;
                    }
                    else if (col == startCol)
                    {
                        cell.CellFormat.VerticalMerge = CellMerge.Previous;
                    }
                    else
                    {
                        cell.CellFormat.HorizontalMerge = CellMerge.Previous;
                        cell.CellFormat.VerticalMerge = CellMerge.Previous;
                    }
                }
            }
        }
    }

    private static MergeCellsParameters ExtractMergeCellsParameters(OperationParameters parameters)
    {
        var startRow = parameters.GetOptional<int?>("startRow");
        var startCol = parameters.GetOptional<int?>("startCol");
        var endRow = parameters.GetOptional<int?>("endRow");
        var endCol = parameters.GetOptional<int?>("endCol");

        if (!startRow.HasValue || !startCol.HasValue || !endRow.HasValue || !endCol.HasValue)
            throw new ArgumentException(
                "startRow, startCol, endRow, and endCol are all required for merge_cells operation");

        var tableIndex = parameters.GetOptional("tableIndex", 0);
        var sectionIndex = parameters.GetOptional<int?>("sectionIndex");

        return new MergeCellsParameters(startRow.Value, startCol.Value, endRow.Value, endCol.Value, tableIndex,
            sectionIndex);
    }

    private sealed record MergeCellsParameters(
        int StartRow,
        int StartCol,
        int EndRow,
        int EndCol,
        int TableIndex,
        int? SectionIndex);
}
