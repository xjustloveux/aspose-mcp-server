using Aspose.Slides;

namespace AsposeMcpServer.Core.ShapeDetailProviders;

/// <summary>
///     Provider for extracting details from Table elements
/// </summary>
public class TableDetailProvider : IShapeDetailProvider
{
    public string TypeName => "Table";

    public bool CanHandle(IShape shape)
    {
        return shape is ITable;
    }

    public object? GetDetails(IShape shape, IPresentation presentation)
    {
        if (shape is not ITable table)
            return null;

        var firstRow = table.FirstRow;

        // Calculate total cells
        var totalCells = table.Rows.Sum(row => row.Count);

        // Get merged cell info
        var mergedCells = new List<object>();
        for (var row = 0; row < table.Rows.Count; row++)
        for (var col = 0; col < table.Columns.Count; col++)
        {
            var cell = table[col, row];
            if (cell.ColSpan > 1 || cell.RowSpan > 1)
                mergedCells.Add(new
                {
                    row,
                    col,
                    colSpan = cell.ColSpan,
                    rowSpan = cell.RowSpan
                });
        }

        return new
        {
            rows = table.Rows.Count,
            columns = table.Columns.Count,
            totalCells,
            firstRow,
            mergedCellCount = mergedCells.Count,
            mergedCells = mergedCells.Count > 0 ? mergedCells : null
        };
    }
}