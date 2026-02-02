using Aspose.Slides;
using AsposeMcpServer.Core.ShapeDetailProviders.Details;

namespace AsposeMcpServer.Core.ShapeDetailProviders.Providers;

/// <summary>
///     Provider for extracting details from Table elements
/// </summary>
public class TableDetailProvider : IShapeDetailProvider
{
    /// <inheritdoc />
    public string TypeName => "Table";

    /// <inheritdoc />
    public bool CanHandle(IShape shape)
    {
        return shape is ITable;
    }

    /// <inheritdoc />
    public ShapeDetails? GetDetails(IShape shape, IPresentation presentation)
    {
        if (shape is not ITable table)
            return null;

        var totalCells = table.Rows.Sum(row => row.Count);

        List<MergedCellInfo> mergedCells = [];
        for (var row = 0; row < table.Rows.Count; row++)
        for (var col = 0; col < table.Columns.Count; col++)
        {
            var cell = table[col, row];
            if (cell.ColSpan > 1 || cell.RowSpan > 1)
                mergedCells.Add(new MergedCellInfo
                {
                    Row = row,
                    Col = col,
                    ColSpan = cell.ColSpan,
                    RowSpan = cell.RowSpan
                });
        }

        var stylePreset = table.StylePreset.ToString();
        if (stylePreset == "None")
            stylePreset = null;

        return new TableDetails
        {
            Rows = table.Rows.Count,
            Columns = table.Columns.Count,
            TotalCells = totalCells,
            FirstRow = table.FirstRow,
            FirstCol = table.FirstCol,
            LastRow = table.LastRow,
            LastCol = table.LastCol,
            StylePreset = stylePreset,
            MergedCellCount = mergedCells.Count,
            MergedCells = mergedCells.Count > 0 ? mergedCells : null
        };
    }
}
