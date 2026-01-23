using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Tables;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Helpers.Word;

/// <summary>
///     Helper methods for Word table operations.
/// </summary>
public static class WordTableHelper
{
    /// <summary>
    ///     Gets tables from a document, optionally filtered by section.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="sectionIndex">The optional section index.</param>
    /// <returns>List of tables.</returns>
    /// <exception cref="ArgumentException">Thrown when section index is out of range.</exception>
    public static List<Table> GetTables(Document doc, int? sectionIndex)
    {
        if (sectionIndex.HasValue)
        {
            if (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count)
                throw new ArgumentException($"Section index {sectionIndex.Value} out of range");
            var section = doc.Sections[sectionIndex.Value];
            return section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        }

        return doc.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
    }

    /// <summary>
    ///     Gets a specific table from the document.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="tableIndex">The table index.</param>
    /// <param name="sectionIndex">The optional section index.</param>
    /// <returns>The table.</returns>
    /// <exception cref="ArgumentException">Thrown when table index is out of range.</exception>
    public static Table GetTable(Document doc, int tableIndex, int? sectionIndex)
    {
        var tables = GetTables(doc, sectionIndex);
        if (tableIndex < 0 || tableIndex >= tables.Count)
            throw new ArgumentException($"Table index {tableIndex} out of range");
        return tables[tableIndex];
    }

    /// <summary>
    ///     Validates section index.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="sectionIndex">The section index.</param>
    /// <exception cref="ArgumentException">Thrown when section index is out of range.</exception>
    public static void ValidateSectionIndex(Document doc, int sectionIndex)
    {
        if (sectionIndex < 0 || sectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
    }

    /// <summary>
    ///     Parses JSON to dictionary of row index to color.
    /// </summary>
    /// <param name="node">The JSON node.</param>
    /// <returns>Dictionary of row colors.</returns>
    public static Dictionary<int, string> ParseColorDictionary(JsonNode? node)
    {
        var result = new Dictionary<int, string>();
        if (node == null) return result;
        try
        {
            var jsonObj = node.AsObject();
            foreach (var kvp in jsonObj)
                if (int.TryParse(kvp.Key, out var key))
                    result[key] = kvp.Value?.GetValue<string>() ?? "";
        }
        catch
        {
            // Ignore parsing errors
        }

        return result;
    }

    /// <summary>
    ///     Parses JSON to list of cell colors.
    /// </summary>
    /// <param name="node">The JSON node.</param>
    /// <returns>List of cell color specifications.</returns>
    public static List<(int row, int col, string color)> ParseCellColors(JsonNode? node)
    {
        List<(int row, int col, string color)> result = [];
        if (node == null) return result;
        try
        {
            var jsonStr = node.ToJsonString();
            var arr = JsonSerializer.Deserialize<JsonElement[][]>(jsonStr);
            if (arr != null)
                foreach (var item in arr)
                    if (item.Length >= 3)
                        result.Add((item[0].GetInt32(), item[1].GetInt32(), item[2].GetString() ?? ""));
        }
        catch
        {
            // Ignore parsing errors
        }

        return result;
    }

    /// <summary>
    ///     Parses JSON to list of merge cell specifications.
    /// </summary>
    /// <param name="node">The JSON node.</param>
    /// <returns>List of merge specifications.</returns>
    public static List<(int startRow, int endRow, int startCol, int endCol)> ParseMergeCells(JsonNode? node)
    {
        List<(int startRow, int endRow, int startCol, int endCol)> result = [];
        if (node == null) return result;
        try
        {
            var jsonStr = node.ToJsonString();
            var arr = JsonSerializer.Deserialize<JsonElement[]>(jsonStr);
            if (arr != null)
                foreach (var item in arr)
                    if (item.TryGetProperty("startRow", out var sr) &&
                        item.TryGetProperty("endRow", out var er) &&
                        item.TryGetProperty("startCol", out var sc) &&
                        item.TryGetProperty("endCol", out var ec))
                        result.Add((sr.GetInt32(), er.GetInt32(), sc.GetInt32(), ec.GetInt32()));
        }
        catch
        {
            // Ignore parsing errors
        }

        return result;
    }

    /// <summary>
    ///     Converts string to CellVerticalAlignment.
    /// </summary>
    /// <param name="alignment">The alignment string.</param>
    /// <returns>The CellVerticalAlignment value.</returns>
    public static CellVerticalAlignment GetVerticalAlignment(string alignment)
    {
        return alignment.ToLowerInvariant() switch
        {
            "top" => CellVerticalAlignment.Top,
            "bottom" => CellVerticalAlignment.Bottom,
            _ => CellVerticalAlignment.Center
        };
    }

    /// <summary>
    ///     Converts string to LineStyle.
    /// </summary>
    /// <param name="style">The style string.</param>
    /// <returns>The LineStyle value.</returns>
    public static LineStyle GetLineStyle(string style)
    {
        return style.ToLowerInvariant() switch
        {
            "none" => LineStyle.None,
            "single" => LineStyle.Single,
            "double" => LineStyle.Double,
            "dotted" => LineStyle.Dot,
            "dashed" => LineStyle.Single,
            "thick" => LineStyle.Thick,
            _ => LineStyle.Single
        };
    }

    /// <summary>
    ///     Gets preceding text for a table.
    /// </summary>
    /// <param name="table">The table.</param>
    /// <param name="maxLength">Maximum text length.</param>
    /// <returns>The preceding text.</returns>
    public static string GetPrecedingText(Table table, int maxLength)
    {
        var precedingSibling = table.PreviousSibling;
        while (precedingSibling != null)
        {
            if (precedingSibling is WordParagraph para)
            {
                var text = para.GetText().Trim();
                if (!string.IsNullOrWhiteSpace(text) && !text.StartsWith('\f'))
                {
                    text = text.Replace("\r", "").Replace("\a", "");
                    if (text.Length > maxLength)
                        return string.Concat(text.AsSpan(0, maxLength), "...");
                    return text;
                }
            }

            precedingSibling = precedingSibling.PreviousSibling;
        }

        return string.Empty;
    }

    /// <summary>
    ///     Applies merge to cells in specified range.
    /// </summary>
    /// <param name="table">The table.</param>
    /// <param name="startRow">Start row index.</param>
    /// <param name="endRow">End row index.</param>
    /// <param name="startCol">Start column index.</param>
    /// <param name="endCol">End column index.</param>
    public static void ApplyMergeCells(Table table, int startRow, int endRow, int startCol,
        int endCol)
    {
        if (!IsValidMergeRange(table, startRow, endRow, startCol, endCol)) return;

        for (var row = startRow; row <= endRow; row++)
        {
            var currentRow = table.Rows[row];
            for (var col = startCol; col <= endCol; col++)
            {
                if (col >= currentRow.Cells.Count) continue;

                var cell = currentRow.Cells[col];
                ApplyCellMerge(cell, row, col, startRow, endRow, startCol, endCol);
            }
        }
    }

    /// <summary>
    ///     Validates if the merge range is valid.
    /// </summary>
    private static bool IsValidMergeRange(Table table, int startRow, int endRow, int startCol,
        int endCol)
    {
        if (startRow > endRow || startCol > endCol) return false;
        if (startRow < 0 || startRow >= table.Rows.Count) return false;
        if (endRow < 0 || endRow >= table.Rows.Count) return false;
        return true;
    }

    /// <summary>
    ///     Applies merge settings to a single cell.
    /// </summary>
    private static void ApplyCellMerge(Cell cell, int row, int col, int startRow, int endRow, int startCol, int endCol)
    {
        var isFirstRow = row == startRow;
        var isFirstCol = col == startCol;
        var hasRowSpan = startRow != endRow;
        var hasColSpan = startCol != endCol;

        if (isFirstRow && isFirstCol)
            ApplyFirstCellMerge(cell, hasRowSpan, hasColSpan);
        else
            ApplySubsequentCellMerge(cell, isFirstRow, isFirstCol);
    }

    /// <summary>
    ///     Applies merge settings to the first cell in the range.
    /// </summary>
    private static void ApplyFirstCellMerge(Cell cell, bool hasRowSpan, bool hasColSpan)
    {
        if (hasRowSpan)
            cell.CellFormat.VerticalMerge = CellMerge.First;
        if (hasColSpan)
            cell.CellFormat.HorizontalMerge = CellMerge.First;
    }

    /// <summary>
    ///     Applies merge settings to subsequent cells in the range.
    /// </summary>
    private static void ApplySubsequentCellMerge(Cell cell, bool isFirstRow, bool isFirstCol)
    {
        if (isFirstRow)
        {
            cell.CellFormat.HorizontalMerge = CellMerge.Previous;
        }
        else if (isFirstCol)
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
