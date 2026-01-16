using Aspose.Words;
using Aspose.Words.Tables;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using WordTable = Aspose.Words.Tables.Table;

namespace AsposeMcpServer.Handlers.Word.Table;

/// <summary>
///     Handler for setting table borders in Word documents.
/// </summary>
public class SetBorderWordTableHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "set_border";

    /// <summary>
    ///     Sets table border properties.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: tableIndex (default 0), sectionIndex, rowIndex, columnIndex,
    ///     borderTop, borderBottom, borderLeft, borderRight, lineStyle (default "single"),
    ///     lineWidth (default 0.5), lineColor (default "000000").
    /// </param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when indices are out of range.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractSetBorderParameters(parameters);

        var doc = context.Document;
        var actualSectionIndex = p.SectionIndex ?? 0;
        if (actualSectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"Section index {actualSectionIndex} out of range");

        var section = doc.Sections[actualSectionIndex];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<WordTable>().ToList();
        if (p.TableIndex >= tables.Count)
            throw new ArgumentException($"Table index {p.TableIndex} out of range");

        var table = tables[p.TableIndex];
        var lineStyleEnum = WordTableHelper.GetLineStyle(p.LineStyle);
        var lineColorParsed = ColorHelper.ParseColor(p.LineColor);

        var targetCells = GetTargetCells(table, p.RowIndex, p.ColumnIndex);

        foreach (var cell in targetCells)
        {
            var borders = cell.CellFormat.Borders;
            if (p.BorderTop)
            {
                borders.Top.LineStyle = lineStyleEnum;
                borders.Top.LineWidth = p.LineWidth;
                borders.Top.Color = lineColorParsed;
            }

            if (p.BorderBottom)
            {
                borders.Bottom.LineStyle = lineStyleEnum;
                borders.Bottom.LineWidth = p.LineWidth;
                borders.Bottom.Color = lineColorParsed;
            }

            if (p.BorderLeft)
            {
                borders.Left.LineStyle = lineStyleEnum;
                borders.Left.LineWidth = p.LineWidth;
                borders.Left.Color = lineColorParsed;
            }

            if (p.BorderRight)
            {
                borders.Right.LineStyle = lineStyleEnum;
                borders.Right.LineWidth = p.LineWidth;
                borders.Right.Color = lineColorParsed;
            }
        }

        MarkModified(context);

        return Success($"Successfully set table {p.TableIndex} borders.");
    }

    /// <summary>
    ///     Gets target cells based on row and column indices.
    /// </summary>
    /// <param name="table">The table.</param>
    /// <param name="rowIndex">Optional row index.</param>
    /// <param name="columnIndex">Optional column index.</param>
    /// <returns>List of target cells.</returns>
    /// <exception cref="ArgumentException">Thrown when indices are out of range.</exception>
    private static List<Cell> GetTargetCells(WordTable table, int? rowIndex, int? columnIndex)
    {
        List<Cell> targetCells = [];

        if (rowIndex.HasValue && columnIndex.HasValue)
        {
            if (rowIndex.Value < table.Rows.Count && columnIndex.Value < table.Rows[rowIndex.Value].Cells.Count)
                targetCells.Add(table.Rows[rowIndex.Value].Cells[columnIndex.Value]);
            else
                throw new ArgumentException($"Row {rowIndex.Value} or column {columnIndex.Value} out of range");
        }
        else if (rowIndex.HasValue)
        {
            if (rowIndex.Value < table.Rows.Count)
                targetCells.AddRange(table.Rows[rowIndex.Value].Cells.Cast<Cell>());
            else
                throw new ArgumentException($"Row {rowIndex.Value} out of range");
        }
        else if (columnIndex.HasValue)
        {
            targetCells.AddRange(table.Rows.Cast<Row>()
                .Where(row => columnIndex.Value < row.Cells.Count)
                .Select(row => row.Cells[columnIndex.Value]));
        }
        else
        {
            targetCells.AddRange(table.Rows.Cast<Row>().SelectMany(row => row.Cells.Cast<Cell>()));
        }

        return targetCells;
    }

    private static SetBorderParameters ExtractSetBorderParameters(OperationParameters parameters)
    {
        var tableIndex = parameters.GetOptional("tableIndex", 0);
        var sectionIndex = parameters.GetOptional<int?>("sectionIndex");
        var rowIndex = parameters.GetOptional<int?>("rowIndex");
        var columnIndex = parameters.GetOptional<int?>("columnIndex");
        var borderTop = parameters.GetOptional("borderTop", false);
        var borderBottom = parameters.GetOptional("borderBottom", false);
        var borderLeft = parameters.GetOptional("borderLeft", false);
        var borderRight = parameters.GetOptional("borderRight", false);
        var lineStyle = parameters.GetOptional("lineStyle", "single");
        var lineWidth = parameters.GetOptional("lineWidth", 0.5);
        var lineColor = parameters.GetOptional("lineColor", "000000");

        return new SetBorderParameters(
            tableIndex,
            sectionIndex,
            rowIndex,
            columnIndex,
            borderTop,
            borderBottom,
            borderLeft,
            borderRight,
            lineStyle,
            lineWidth,
            lineColor);
    }

    private record SetBorderParameters(
        int TableIndex,
        int? SectionIndex,
        int? RowIndex,
        int? ColumnIndex,
        bool BorderTop,
        bool BorderBottom,
        bool BorderLeft,
        bool BorderRight,
        string LineStyle,
        double LineWidth,
        string LineColor);
}
