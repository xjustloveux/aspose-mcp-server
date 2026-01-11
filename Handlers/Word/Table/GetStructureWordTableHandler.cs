using System.Text;
using Aspose.Words;
using Aspose.Words.Tables;
using AsposeMcpServer.Core.Handlers;
using WordTable = Aspose.Words.Tables.Table;

namespace AsposeMcpServer.Handlers.Word.Table;

/// <summary>
///     Handler for getting table structure in Word documents.
/// </summary>
public class GetStructureWordTableHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get_structure";

    /// <summary>
    ///     Gets detailed table structure information.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: tableIndex (default 0), sectionIndex, includeContent (default false),
    ///     includeCellFormatting (default true).
    /// </param>
    /// <returns>Formatted string containing table structure information.</returns>
    /// <exception cref="ArgumentException">Thrown when tableIndex or sectionIndex is out of range.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var tableIndex = parameters.GetOptional("tableIndex", 0);
        var sectionIndex = parameters.GetOptional<int?>("sectionIndex");
        var includeContent = parameters.GetOptional("includeContent", false);
        var includeCellFormatting = parameters.GetOptional("includeCellFormatting", true);

        var doc = context.Document;
        var actualSectionIndex = sectionIndex ?? 0;
        if (actualSectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"Section index {actualSectionIndex} out of range");

        var section = doc.Sections[actualSectionIndex];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<WordTable>().ToList();
        if (tableIndex < 0 || tableIndex >= tables.Count)
            throw new ArgumentException($"Table index {tableIndex} out of range");

        var table = tables[tableIndex];
        var result = new StringBuilder();

        AppendBasicInfo(result, tableIndex, table);
        AppendTableFormat(result, table);

        if (includeContent)
            AppendContentPreview(result, table);

        if (includeCellFormatting && table.Rows.Count > 0 && table.Rows[0].Cells.Count > 0)
            AppendCellFormatting(result, table);

        return result.ToString();
    }

    /// <summary>
    ///     Appends basic table info to the result.
    /// </summary>
    /// <param name="result">The result builder.</param>
    /// <param name="tableIndex">The table index.</param>
    /// <param name="table">The table.</param>
    private static void AppendBasicInfo(StringBuilder result, int tableIndex, WordTable table)
    {
        result.AppendLine($"=== Table #{tableIndex} Structure ===\n");
        result.AppendLine("[Basic Info]");
        result.AppendLine($"Rows: {table.Rows.Count}");
        if (table.Rows.Count > 0)
            result.AppendLine($"Columns: {table.Rows[0].Cells.Count}");
        result.AppendLine();
    }

    /// <summary>
    ///     Appends table format info to the result.
    /// </summary>
    /// <param name="result">The result builder.</param>
    /// <param name="table">The table.</param>
    private static void AppendTableFormat(StringBuilder result, WordTable table)
    {
        result.AppendLine("[Table Format]");
        result.AppendLine($"Alignment: {table.Alignment}");
        result.AppendLine($"Style: {table.Style?.Name ?? "None"}");
        result.AppendLine($"Left Indent: {table.LeftIndent:F2} pt");
        if (table.PreferredWidth.Type != PreferredWidthType.Auto)
            result.AppendLine($"Width: {table.PreferredWidth.Value} ({table.PreferredWidth.Type})");
        result.AppendLine($"Allow Auto Fit: {table.AllowAutoFit}");
        result.AppendLine();
    }

    /// <summary>
    ///     Appends content preview to the result.
    /// </summary>
    /// <param name="result">The result builder.</param>
    /// <param name="table">The table.</param>
    private static void AppendContentPreview(StringBuilder result, WordTable table)
    {
        result.AppendLine("[Content Preview]");
        for (var i = 0; i < Math.Min(table.Rows.Count, 5); i++)
        {
            var row = table.Rows[i];
            result.Append($"  Row {i}: | ");
            for (var j = 0; j < row.Cells.Count; j++)
            {
                var cell = row.Cells[j];
                var cellText = cell.GetText().Trim().Replace("\r", "").Replace("\n", " ");
                if (cellText.Length > 30)
                    cellText = cellText.Substring(0, 27) + "...";
                result.Append($"{cellText} | ");
            }

            result.AppendLine();
        }

        if (table.Rows.Count > 5)
            result.AppendLine($"  ... ({table.Rows.Count - 5} more rows)");
        result.AppendLine();
    }

    /// <summary>
    ///     Appends first cell formatting info to the result.
    /// </summary>
    /// <param name="result">The result builder.</param>
    /// <param name="table">The table.</param>
    private static void AppendCellFormatting(StringBuilder result, WordTable table)
    {
        result.AppendLine("[First Cell Formatting]");
        var cell = table.Rows[0].Cells[0];
        result.AppendLine($"Top Padding: {cell.CellFormat.TopPadding:F2} pt");
        result.AppendLine($"Bottom Padding: {cell.CellFormat.BottomPadding:F2} pt");
        result.AppendLine($"Left Padding: {cell.CellFormat.LeftPadding:F2} pt");
        result.AppendLine($"Right Padding: {cell.CellFormat.RightPadding:F2} pt");
        result.AppendLine($"Vertical Alignment: {cell.CellFormat.VerticalAlignment}");
        result.AppendLine();
    }
}
