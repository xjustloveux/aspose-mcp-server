using Aspose.Cells;

namespace AsposeMcpServer.Helpers.Excel;

/// <summary>
///     Helper class for Excel hyperlink operations.
/// </summary>
public static class ExcelHyperlinkHelper
{
    /// <summary>
    ///     Finds hyperlink index by cell reference.
    /// </summary>
    /// <param name="hyperlinks">The hyperlink collection to search.</param>
    /// <param name="cell">The cell reference in A1 notation.</param>
    /// <returns>The hyperlink index if found, otherwise null.</returns>
    public static int? FindHyperlinkIndexByCell(HyperlinkCollection hyperlinks, string cell)
    {
        CellsHelper.CellNameToIndex(cell, out var rowIndex, out var colIndex);

        for (var i = 0; i < hyperlinks.Count; i++)
        {
            var area = hyperlinks[i].Area;
            if (rowIndex >= area.StartRow && rowIndex <= area.EndRow &&
                colIndex >= area.StartColumn && colIndex <= area.EndColumn)
                return i;
        }

        return null;
    }

    /// <summary>
    ///     Resolves hyperlink index from either direct index or cell reference.
    /// </summary>
    /// <param name="hyperlinks">The hyperlink collection to search.</param>
    /// <param name="hyperlinkIndex">The direct hyperlink index, or null to use cell reference.</param>
    /// <param name="cell">The cell reference in A1 notation as an alternative to index.</param>
    /// <returns>The resolved hyperlink index.</returns>
    /// <exception cref="ArgumentException">Thrown when neither index nor cell is provided, or hyperlink is not found.</exception>
    public static int ResolveHyperlinkIndex(HyperlinkCollection hyperlinks, int? hyperlinkIndex, string? cell)
    {
        int index;

        if (hyperlinkIndex.HasValue)
        {
            index = hyperlinkIndex.Value;
        }
        else if (!string.IsNullOrEmpty(cell))
        {
            var foundIndex = FindHyperlinkIndexByCell(hyperlinks, cell);
            if (!foundIndex.HasValue)
                throw new ArgumentException($"No hyperlink found at cell {cell}.");
            index = foundIndex.Value;
        }
        else
        {
            throw new ArgumentException("Either 'hyperlinkIndex' or 'cell' must be provided.");
        }

        if (index < 0 || index >= hyperlinks.Count)
            throw new ArgumentException(
                $"Hyperlink index {index} is out of range. Worksheet has {hyperlinks.Count} hyperlinks.");

        return index;
    }
}
