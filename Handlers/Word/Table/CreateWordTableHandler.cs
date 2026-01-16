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
        var tableParams = ExtractTableParameters(parameters);
        var doc = context.Document;
        var builder = new DocumentBuilder(doc);

        var (section, actualSectionIndex) = ValidateAndGetSection(doc, tableParams.SectionIndex);
        var parsedTableData = ParseTableData(tableParams.TableData);
        var (numRows, numCols) = CalculateTableDimensions(parsedTableData, tableParams.Rows, tableParams.Columns);

        MoveToInsertPosition(builder, section, tableParams.ParagraphIndex, actualSectionIndex);
        var table = BuildTable(builder, numRows, numCols, parsedTableData, tableParams);
        FinalizeTable(table, tableParams);

        MarkModified(context);
        return Success($"Successfully created table with {numRows} rows and {numCols} columns.");
    }

    /// <summary>
    ///     Extracts table creation parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted table creation parameters.</returns>
    private static TableCreationParameters ExtractTableParameters(OperationParameters parameters)
    {
        return new TableCreationParameters(
            parameters.GetOptional("paragraphIndex", -1),
            parameters.GetOptional<int?>("rows"),
            parameters.GetOptional<int?>("columns"),
            parameters.GetOptional<string?>("tableData"),
            parameters.GetOptional<double?>("tableWidth"),
            parameters.GetOptional("autoFit", true),
            parameters.GetOptional("hasHeader", true),
            parameters.GetOptional<string?>("headerBackgroundColor"),
            parameters.GetOptional<string?>("cellBackgroundColor"),
            parameters.GetOptional<string?>("alternatingRowColor"),
            parameters.GetOptional<string?>("rowColors"),
            parameters.GetOptional<string?>("cellColors"),
            parameters.GetOptional<string?>("mergeCells"),
            parameters.GetOptional<string?>("fontName"),
            parameters.GetOptional<double?>("fontSize"),
            parameters.GetOptional("verticalAlignment", "center"),
            parameters.GetOptional<int?>("sectionIndex")
        );
    }

    /// <summary>
    ///     Validates and gets the target section.
    /// </summary>
    /// <param name="doc">The Word document.</param>
    /// <param name="sectionIndex">The section index.</param>
    /// <returns>A tuple containing the section and actual index.</returns>
    /// <exception cref="ArgumentException">Thrown when section index is out of range.</exception>
    private static (Section section, int actualIndex) ValidateAndGetSection(Document doc, int? sectionIndex)
    {
        var actualSectionIndex = sectionIndex ?? 0;
        if (actualSectionIndex < 0 || actualSectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
        return (doc.Sections[actualSectionIndex], actualSectionIndex);
    }

    /// <summary>
    ///     Calculates the table dimensions.
    /// </summary>
    /// <param name="parsedData">The parsed table data.</param>
    /// <param name="rows">The specified number of rows.</param>
    /// <param name="columns">The specified number of columns.</param>
    /// <returns>A tuple containing the number of rows and columns.</returns>
    private static (int numRows, int numCols) CalculateTableDimensions(List<List<string>>? parsedData, int? rows,
        int? columns)
    {
        if (parsedData is { Count: > 0 })
            return (parsedData.Count, parsedData.Max(r => r.Count));
        return (rows ?? 3, columns ?? 3);
    }

    /// <summary>
    ///     Moves the builder to the insert position.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="section">The document section.</param>
    /// <param name="paragraphIndex">The paragraph index.</param>
    /// <param name="sectionIndex">The section index.</param>
    private static void MoveToInsertPosition(DocumentBuilder builder, Section section, int paragraphIndex,
        int sectionIndex)
    {
        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
        if (paragraphIndex >= 0 && paragraphIndex < paragraphs.Count)
            builder.MoveTo(paragraphs[paragraphIndex]);
        else
            builder.MoveToSection(sectionIndex);
    }

    /// <summary>
    ///     Builds the table with the specified dimensions and data.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="numRows">The number of rows.</param>
    /// <param name="numCols">The number of columns.</param>
    /// <param name="parsedTableData">The parsed table data.</param>
    /// <param name="p">The table creation parameters.</param>
    /// <returns>The created table.</returns>
    private static Aspose.Words.Tables.Table BuildTable(DocumentBuilder builder, int numRows, int numCols,
        List<List<string>>? parsedTableData, TableCreationParameters p)
    {
        var colorContext = CreateColorContext(p);
        var table = builder.StartTable();

        for (var i = 0; i < numRows; i++)
        {
            for (var j = 0; j < numCols; j++)
                InsertTableCell(builder, i, j, parsedTableData, p, colorContext);
            builder.EndRow();
        }

        builder.EndTable();
        return table;
    }

    /// <summary>
    ///     Creates the color context from parameters.
    /// </summary>
    /// <param name="p">The table creation parameters.</param>
    /// <returns>The color context.</returns>
    private static ColorContext CreateColorContext(TableCreationParameters p)
    {
        return new ColorContext(
            !string.IsNullOrEmpty(p.RowColors) ? WordTableHelper.ParseColorDictionary(JsonNode.Parse(p.RowColors)) : [],
            !string.IsNullOrEmpty(p.CellColors) ? WordTableHelper.ParseCellColors(JsonNode.Parse(p.CellColors)) : []
        );
    }

    /// <summary>
    ///     Inserts a cell into the table.
    /// </summary>
    /// <param name="builder">The document builder.</param>
    /// <param name="row">The row index.</param>
    /// <param name="col">The column index.</param>
    /// <param name="parsedTableData">The parsed table data.</param>
    /// <param name="p">The table creation parameters.</param>
    /// <param name="colorContext">The color context.</param>
    private static void InsertTableCell(DocumentBuilder builder, int row, int col,
        List<List<string>>? parsedTableData, TableCreationParameters p, ColorContext colorContext)
    {
        builder.InsertCell();

        if (builder.CurrentParagraph.ParentNode is Cell cell)
        {
            cell.CellFormat.VerticalAlignment = WordTableHelper.GetVerticalAlignment(p.VerticalAlignment);
            ApplyCellBackgroundColor(cell, row, col, colorContext.CellColors, colorContext.RowColors, p.HasHeader,
                p.HeaderBackgroundColor, p.AlternatingRowColor, p.CellBackgroundColor);
        }

        var cellText = parsedTableData != null && row < parsedTableData.Count && col < parsedTableData[row].Count
            ? parsedTableData[row][col]
            : "";

        if (p.FontSize.HasValue) builder.Font.Size = p.FontSize.Value;
        if (!string.IsNullOrEmpty(p.FontName)) builder.Font.Name = p.FontName;
        builder.Font.Bold = p.HasHeader && row == 0;

        WriteCellText(builder, cellText);
    }

    /// <summary>
    ///     Finalizes the table with width, auto-fit, and merge settings.
    /// </summary>
    /// <param name="table">The table to finalize.</param>
    /// <param name="p">The table creation parameters.</param>
    private static void FinalizeTable(Aspose.Words.Tables.Table table, TableCreationParameters p)
    {
        if (p.TableWidth.HasValue)
            table.PreferredWidth = PreferredWidth.FromPoints(p.TableWidth.Value);
        table.AllowAutoFit = p.AutoFit;

        var mergeCellsList = !string.IsNullOrEmpty(p.MergeCells)
            ? WordTableHelper.ParseMergeCells(JsonNode.Parse(p.MergeCells))
            : [];

        foreach (var merge in mergeCellsList)
            WordTableHelper.ApplyMergeCells(table, merge.startRow, merge.endRow, merge.startCol, merge.endCol);
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

    /// <summary>
    ///     Record to hold table creation parameters.
    /// </summary>
    private record TableCreationParameters(
        int ParagraphIndex,
        int? Rows,
        int? Columns,
        string? TableData,
        double? TableWidth,
        bool AutoFit,
        bool HasHeader,
        string? HeaderBackgroundColor,
        string? CellBackgroundColor,
        string? AlternatingRowColor,
        string? RowColors,
        string? CellColors,
        string? MergeCells,
        string? FontName,
        double? FontSize,
        string VerticalAlignment,
        int? SectionIndex);

    /// <summary>
    ///     Record to hold color context for table cells.
    /// </summary>
    private record ColorContext(
        Dictionary<int, string> RowColors,
        List<(int row, int col, string color)> CellColors);
}
