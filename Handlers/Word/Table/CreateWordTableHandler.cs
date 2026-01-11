using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Tables;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Table;

/// <summary>
///     Handler for creating tables in Word documents.
/// </summary>
public class CreateWordTableHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "create";

    /// <summary>
    ///     Creates a new table in the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: rows, columns, paragraphIndex, tableData, tableWidth, autoFit, hasHeader,
    ///     headerBackgroundColor, cellBackgroundColor, alternatingRowColor, rowColors, cellColors,
    ///     mergeCells, fontName, fontSize, verticalAlignment, sectionIndex.
    /// </param>
    /// <returns>Success message with table details.</returns>
    /// <exception cref="ArgumentException">Thrown when sectionIndex is out of range or tableData JSON is invalid.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var paragraphIndex = parameters.GetOptional("paragraphIndex", -1);
        var rows = parameters.GetOptional<int?>("rows");
        var columns = parameters.GetOptional<int?>("columns");
        var tableData = parameters.GetOptional<string?>("tableData");
        var tableWidth = parameters.GetOptional<double?>("tableWidth");
        var autoFit = parameters.GetOptional("autoFit", true);
        var hasHeader = parameters.GetOptional("hasHeader", true);
        var headerBackgroundColor = parameters.GetOptional<string?>("headerBackgroundColor");
        var cellBackgroundColor = parameters.GetOptional<string?>("cellBackgroundColor");
        var alternatingRowColor = parameters.GetOptional<string?>("alternatingRowColor");
        var rowColors = parameters.GetOptional<string?>("rowColors");
        var cellColors = parameters.GetOptional<string?>("cellColors");
        var mergeCells = parameters.GetOptional<string?>("mergeCells");
        var fontName = parameters.GetOptional<string?>("fontName");
        var fontSize = parameters.GetOptional<double?>("fontSize");
        var verticalAlignment = parameters.GetOptional("verticalAlignment", "center");
        var sectionIndex = parameters.GetOptional<int?>("sectionIndex");

        var doc = context.Document;
        var builder = new DocumentBuilder(doc);

        var actualSectionIndex = sectionIndex ?? 0;
        if (actualSectionIndex < 0 || actualSectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

        var section = doc.Sections[actualSectionIndex];

        var parsedTableData = ParseTableData(tableData);

        int numRows, numCols;
        if (parsedTableData is { Count: > 0 })
        {
            numRows = parsedTableData.Count;
            numCols = parsedTableData.Max(r => r.Count);
        }
        else
        {
            numRows = rows ?? 3;
            numCols = columns ?? 3;
        }

        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();

        if (paragraphIndex >= 0 && paragraphIndex < paragraphs.Count)
            builder.MoveTo(paragraphs[paragraphIndex]);
        else
            builder.MoveToSection(actualSectionIndex);

        var table = builder.StartTable();

        var rowColorsDict = !string.IsNullOrEmpty(rowColors)
            ? WordTableHelper.ParseColorDictionary(JsonNode.Parse(rowColors))
            : new Dictionary<int, string>();
        var cellColorsList = !string.IsNullOrEmpty(cellColors)
            ? WordTableHelper.ParseCellColors(JsonNode.Parse(cellColors))
            : [];
        var mergeCellsList = !string.IsNullOrEmpty(mergeCells)
            ? WordTableHelper.ParseMergeCells(JsonNode.Parse(mergeCells))
            : [];

        for (var i = 0; i < numRows; i++)
        {
            for (var j = 0; j < numCols; j++)
            {
                builder.InsertCell();

                if (builder.CurrentParagraph.ParentNode is Cell cell)
                {
                    cell.CellFormat.VerticalAlignment = WordTableHelper.GetVerticalAlignment(verticalAlignment);
                    ApplyCellBackgroundColor(cell, i, j, cellColorsList, rowColorsDict, hasHeader,
                        headerBackgroundColor, alternatingRowColor, cellBackgroundColor);
                }

                var cellText = "";
                if (parsedTableData != null && i < parsedTableData.Count && j < parsedTableData[i].Count)
                    cellText = parsedTableData[i][j];

                if (fontSize.HasValue)
                    builder.Font.Size = fontSize.Value;

                if (!string.IsNullOrEmpty(fontName))
                    builder.Font.Name = fontName;

                builder.Font.Bold = hasHeader && i == 0;

                WriteCellText(builder, cellText);
            }

            builder.EndRow();
        }

        builder.EndTable();

        if (tableWidth.HasValue)
            table.PreferredWidth = PreferredWidth.FromPoints(tableWidth.Value);

        table.AllowAutoFit = autoFit;

        foreach (var merge in mergeCellsList)
            WordTableHelper.ApplyMergeCells(table, merge.startRow, merge.endRow, merge.startCol, merge.endCol);

        MarkModified(context);

        return Success($"Successfully created table with {numRows} rows and {numCols} columns.");
    }

    /// <summary>
    ///     Parses table data from JSON string.
    /// </summary>
    /// <param name="tableData">JSON string representing 2D array of table data.</param>
    /// <returns>Parsed table data as list of lists.</returns>
    /// <exception cref="ArgumentException">Thrown when JSON format is invalid.</exception>
    private static List<List<string>>? ParseTableData(string? tableData)
    {
        if (string.IsNullOrEmpty(tableData)) return null;

        try
        {
            var jsonArray = JsonSerializer.Deserialize<JsonElement>(tableData);
            if (jsonArray.ValueKind != JsonValueKind.Array) return null;

            List<List<string>> result = [];
            foreach (var row in jsonArray.EnumerateArray())
            {
                List<string> rowList = [];
                foreach (var cell in row.EnumerateArray())
                    rowList.Add(cell.ToString());
                result.Add(rowList);
            }

            return result;
        }
        catch (Exception ex)
        {
            throw new ArgumentException($"Invalid tableData JSON format: {ex.Message}");
        }
    }

    /// <summary>
    ///     Applies background color to a cell based on various color settings.
    /// </summary>
    /// <param name="cell">The cell to apply color to.</param>
    /// <param name="rowIndex">The row index.</param>
    /// <param name="colIndex">The column index.</param>
    /// <param name="cellColorsList">List of specific cell colors.</param>
    /// <param name="rowColorsDict">Dictionary of row colors.</param>
    /// <param name="hasHeader">Whether table has header row.</param>
    /// <param name="headerBackgroundColor">Header background color.</param>
    /// <param name="alternatingRowColor">Alternating row color.</param>
    /// <param name="cellBackgroundColor">Default cell background color.</param>
    private static void ApplyCellBackgroundColor(Cell cell, int rowIndex, int colIndex,
        List<(int row, int col, string color)> cellColorsList, Dictionary<int, string> rowColorsDict,
        bool hasHeader, string? headerBackgroundColor, string? alternatingRowColor, string? cellBackgroundColor)
    {
        var specificColor = cellColorsList.FirstOrDefault(c => c.row == rowIndex && c.col == colIndex);
        if (!string.IsNullOrEmpty(specificColor.color))
            cell.CellFormat.Shading.BackgroundPatternColor = ColorHelper.ParseColor(specificColor.color, true);
        else if (rowColorsDict.TryGetValue(rowIndex, out var rowColor))
            cell.CellFormat.Shading.BackgroundPatternColor = ColorHelper.ParseColor(rowColor, true);
        else if (hasHeader && rowIndex == 0 && !string.IsNullOrEmpty(headerBackgroundColor))
            cell.CellFormat.Shading.BackgroundPatternColor = ColorHelper.ParseColor(headerBackgroundColor, true);
        else if (!string.IsNullOrEmpty(alternatingRowColor) && rowIndex > 0 && rowIndex % 2 == 0)
            cell.CellFormat.Shading.BackgroundPatternColor = ColorHelper.ParseColor(alternatingRowColor, true);
        else if (!string.IsNullOrEmpty(cellBackgroundColor))
            cell.CellFormat.Shading.BackgroundPatternColor = ColorHelper.ParseColor(cellBackgroundColor, true);
    }

    /// <summary>
    ///     Writes text to a cell, handling line breaks.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="cellText">The text to write.</param>
    private static void WriteCellText(DocumentBuilder builder, string cellText)
    {
        if (string.IsNullOrEmpty(cellText)) return;

        if (cellText.Contains('\n'))
        {
            var lines = cellText.Split('\n');
            for (var lineIdx = 0; lineIdx < lines.Length; lineIdx++)
            {
                if (lineIdx > 0)
                    builder.InsertBreak(BreakType.LineBreak);
                builder.Write(lines[lineIdx]);
            }
        }
        else
        {
            builder.Write(cellText);
        }
    }
}
