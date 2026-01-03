using System.ComponentModel;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Tables;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for managing tables in Word documents
///     Merges all table-related operations into a single tool
/// </summary>
[McpServerToolType]
public class WordTableTool
{
    /// <summary>
    ///     Session manager for document session operations
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the WordTableTool class
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations</param>
    public WordTableTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "word_table")]
    [Description(
        @"Manage tables in Word documents. Supports 16 operations: create, delete, get, insert_row, delete_row, insert_column, delete_column, merge_cells, split_cell, edit_cell_format, move_table, copy_table, get_structure, set_border, set_column_width, set_row_height.

Usage examples:
- Create table: word_table(operation='create', path='doc.docx', rows=3, columns=4)
- Delete table: word_table(operation='delete', path='doc.docx', tableIndex=0)
- Get tables: word_table(operation='get', path='doc.docx')
- Insert row: word_table(operation='insert_row', path='doc.docx', tableIndex=0, rowIndex=1)
- Merge cells: word_table(operation='merge_cells', path='doc.docx', tableIndex=0, startRow=0, startCol=0, endRow=1, endCol=1)
- Set border: word_table(operation='set_border', path='doc.docx', tableIndex=0, borderTop=true, borderBottom=true)

Notes:
- All indices are 0-based
- Use rowData/columnData as JSON arrays to provide data when inserting rows/columns
- Use sectionIndex to specify which section's tables to work with
- cellColors format: [[row, col, '#RRGGBB'], ...] for per-cell coloring
- mergeCells format: [{startRow, endRow, startCol, endCol}, ...] for batch merging")]
    public string Execute(
        [Description(
            "Operation: create, delete, get, insert_row, delete_row, insert_column, delete_column, merge_cells, split_cell, edit_cell_format, move_table, copy_table, get_structure, set_border, set_column_width, set_row_height")]
        string operation,
        [Description("Document file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Table index (0-based)")] int tableIndex = 0,
        [Description("Section index (0-based)")]
        int? sectionIndex = null,
        [Description("Number of rows (for create)")]
        int? rows = null,
        [Description("Number of columns (for create)")]
        int? columns = null,
        [Description("Paragraph index to insert after (-1 for end, for create)")]
        int paragraphIndex = -1,
        [Description("Table data as JSON 2D array (for create)")]
        string? tableData = null,
        [Description("Table width in points (for create)")]
        double? tableWidth = null,
        [Description("Auto-fit table (for create, default: true)")]
        bool autoFit = true,
        [Description("Header row with alternating colors (for create, default: true)")]
        bool hasHeader = true,
        [Description("Header background color hex (for create)")]
        string? headerBackgroundColor = null,
        [Description("Cell background color hex (for create)")]
        string? cellBackgroundColor = null,
        [Description("Alternating row color hex (for create)")]
        string? alternatingRowColor = null,
        [Description("Row colors by index as JSON object (for create)")]
        string? rowColors = null,
        [Description("Cell colors as JSON array (for create)")]
        string? cellColors = null,
        [Description("Cells to merge as JSON array (for create)")]
        string? mergeCells = null,
        [Description("Font name (for create)")]
        string? fontName = null,
        [Description("Font size in points (for create)")]
        double? fontSize = null,
        [Description("Vertical alignment: top, center, bottom (for create, default: center)")]
        string verticalAlignment = "center",
        [Description("Row index (0-based, for row/cell operations)")]
        int? rowIndex = null,
        [Description("Column index (0-based, for column/cell operations)")]
        int? columnIndex = null,
        [Description("Row data as JSON array (for insert_row)")]
        string? rowData = null,
        [Description("Column data as JSON array (for insert_column)")]
        string? columnData = null,
        [Description("Insert before target position (default: false)")]
        bool insertBefore = false,
        [Description("Start row for merge (0-based)")]
        int? startRow = null,
        [Description("End row for merge (0-based)")]
        int? endRow = null,
        [Description("Start column for merge (0-based)")]
        int? startCol = null,
        [Description("End column for merge (0-based)")]
        int? endCol = null,
        [Description("Number of rows to split into (for split_cell, default: 2)")]
        int splitRows = 2,
        [Description("Number of columns to split into (for split_cell, default: 2)")]
        int splitCols = 2,
        [Description("Apply formatting to entire row (for edit_cell_format)")]
        bool applyToRow = false,
        [Description("Apply formatting to entire column (for edit_cell_format)")]
        bool applyToColumn = false,
        [Description("Apply formatting to entire table (for edit_cell_format)")]
        bool applyToTable = false,
        [Description("Background color hex (for edit_cell_format)")]
        string? backgroundColor = null,
        [Description("Text alignment: left, center, right, justify (for edit_cell_format)")]
        string? alignment = null,
        [Description("Vertical alignment: top, center, bottom (for edit_cell_format)")]
        string? verticalAlignmentFormat = null,
        [Description("Top padding in points (for edit_cell_format)")]
        double? paddingTop = null,
        [Description("Bottom padding in points (for edit_cell_format)")]
        double? paddingBottom = null,
        [Description("Left padding in points (for edit_cell_format)")]
        double? paddingLeft = null,
        [Description("Right padding in points (for edit_cell_format)")]
        double? paddingRight = null,
        [Description("Font name for ASCII (for edit_cell_format)")]
        string? fontNameAscii = null,
        [Description("Font name for Far East (for edit_cell_format)")]
        string? fontNameFarEast = null,
        [Description("Font size for cells in points (for edit_cell_format)")]
        double? cellFontSize = null,
        [Description("Bold text (for edit_cell_format)")]
        bool? bold = null,
        [Description("Italic text (for edit_cell_format)")]
        bool? italic = null,
        [Description("Text color hex (for edit_cell_format)")]
        string? color = null,
        [Description("Target paragraph index for move/copy (0-based, -1 for end)")]
        int targetParagraphIndex = -1,
        [Description("Source section index for copy (0-based)")]
        int? sourceSectionIndex = null,
        [Description("Target section index for copy (0-based)")]
        int? targetSectionIndex = null,
        [Description("Include content in get_structure (default: false)")]
        bool includeContent = false,
        [Description("Include cell formatting in get_structure (default: true)")]
        bool includeCellFormatting = true,
        [Description("Enable top border (for set_border)")]
        bool borderTop = false,
        [Description("Enable bottom border (for set_border)")]
        bool borderBottom = false,
        [Description("Enable left border (for set_border)")]
        bool borderLeft = false,
        [Description("Enable right border (for set_border)")]
        bool borderRight = false,
        [Description(
            "Border line style: none, single, double, dotted, dashed, thick (for set_border, default: single)")]
        string lineStyle = "single",
        [Description("Border line width in points (for set_border, default: 0.5)")]
        double lineWidth = 0.5,
        [Description("Border line color hex (for set_border, default: 000000)")]
        string lineColor = "000000",
        [Description("Column width in points (for set_column_width)")]
        double? columnWidth = null,
        [Description("Row height in points (for set_row_height)")]
        double? rowHeight = null,
        [Description("Height rule: auto, atLeast, exactly (for set_row_height, default: atLeast)")]
        string heightRule = "atLeast")
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);

        return operation.ToLower() switch
        {
            "create" => CreateTable(ctx, outputPath, paragraphIndex, rows, columns, tableData, tableWidth, autoFit,
                hasHeader, headerBackgroundColor, cellBackgroundColor, alternatingRowColor, rowColors, cellColors,
                mergeCells, fontName, fontSize, verticalAlignment, sectionIndex),
            "delete" => DeleteTable(ctx, outputPath, tableIndex, sectionIndex),
            "get" => GetTables(ctx, sectionIndex),
            "insert_row" => InsertRow(ctx, outputPath, tableIndex, rowIndex, rowData, insertBefore, sectionIndex),
            "delete_row" => DeleteRow(ctx, outputPath, tableIndex, rowIndex, sectionIndex),
            "insert_column" => InsertColumn(ctx, outputPath, tableIndex, columnIndex, columnData, insertBefore,
                sectionIndex),
            "delete_column" => DeleteColumn(ctx, outputPath, tableIndex, columnIndex, sectionIndex),
            "merge_cells" => MergeCells(ctx, outputPath, tableIndex, startRow, startCol, endRow, endCol, sectionIndex),
            "split_cell" => SplitCell(ctx, outputPath, tableIndex, rowIndex, columnIndex, splitRows, splitCols,
                sectionIndex),
            "edit_cell_format" => EditCellFormat(ctx, outputPath, tableIndex, rowIndex, columnIndex, applyToRow,
                applyToColumn, applyToTable, backgroundColor, alignment, verticalAlignmentFormat, paddingTop,
                paddingBottom, paddingLeft, paddingRight, fontName, fontNameAscii, fontNameFarEast, cellFontSize, bold,
                italic, color, sectionIndex),
            "move_table" => MoveTable(ctx, outputPath, tableIndex, targetParagraphIndex, sectionIndex),
            "copy_table" => CopyTable(ctx, outputPath, tableIndex, targetParagraphIndex, sourceSectionIndex,
                targetSectionIndex),
            "get_structure" => GetTableStructure(ctx, tableIndex, sectionIndex, includeContent, includeCellFormatting),
            "set_border" => SetTableBorder(ctx, outputPath, tableIndex, sectionIndex, rowIndex, columnIndex, borderTop,
                borderBottom, borderLeft, borderRight, lineStyle, lineWidth, lineColor),
            "set_column_width" => SetColumnWidth(ctx, outputPath, tableIndex, columnIndex, columnWidth, sectionIndex),
            "set_row_height" => SetRowHeight(ctx, outputPath, tableIndex, rowIndex, rowHeight, heightRule,
                sectionIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Creates a new table in the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="paragraphIndex">The paragraph index to insert after.</param>
    /// <param name="rows">The number of rows.</param>
    /// <param name="columns">The number of columns.</param>
    /// <param name="tableData">The table data as JSON 2D array.</param>
    /// <param name="tableWidth">The table width in points.</param>
    /// <param name="autoFit">Whether to auto-fit the table.</param>
    /// <param name="hasHeader">Whether to have a header row.</param>
    /// <param name="headerBackgroundColor">The header background color in hex format.</param>
    /// <param name="cellBackgroundColor">The cell background color in hex format.</param>
    /// <param name="alternatingRowColor">The alternating row color in hex format.</param>
    /// <param name="rowColors">The row colors by index as JSON object.</param>
    /// <param name="cellColors">The cell colors as JSON array.</param>
    /// <param name="mergeCells">The cells to merge as JSON array.</param>
    /// <param name="fontName">The font name.</param>
    /// <param name="fontSize">The font size in points.</param>
    /// <param name="verticalAlignment">The vertical alignment (top, center, bottom).</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when sectionIndex is out of range or tableData JSON is invalid.</exception>
    private static string CreateTable(DocumentContext<Document> ctx, string? outputPath, int paragraphIndex, int? rows,
        int? columns, string? tableData, double? tableWidth, bool autoFit, bool hasHeader,
        string? headerBackgroundColor, string? cellBackgroundColor, string? alternatingRowColor, string? rowColors,
        string? cellColors, string? mergeCells, string? fontName, double? fontSize, string verticalAlignment,
        int? sectionIndex)
    {
        var doc = ctx.Document;
        var builder = new DocumentBuilder(doc);

        var actualSectionIndex = sectionIndex ?? 0;
        if (actualSectionIndex < 0 || actualSectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

        var section = doc.Sections[actualSectionIndex];

        List<List<string>>? parsedTableData = null;
        if (!string.IsNullOrEmpty(tableData))
            try
            {
                var jsonArray = JsonSerializer.Deserialize<JsonElement>(tableData);
                if (jsonArray.ValueKind == JsonValueKind.Array)
                {
                    parsedTableData = [];
                    foreach (var row in jsonArray.EnumerateArray())
                    {
                        List<string> rowList = [];
                        foreach (var cell in row.EnumerateArray())
                            rowList.Add(cell.ToString());
                        parsedTableData.Add(rowList);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"Invalid tableData JSON format: {ex.Message}");
            }

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

        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        if (paragraphIndex >= 0 && paragraphIndex < paragraphs.Count)
            builder.MoveTo(paragraphs[paragraphIndex]);
        else
            builder.MoveToSection(actualSectionIndex);

        var table = builder.StartTable();

        var rowColorsDict = !string.IsNullOrEmpty(rowColors)
            ? ParseColorDictionary(JsonNode.Parse(rowColors))
            : new Dictionary<int, string>();
        var cellColorsList = !string.IsNullOrEmpty(cellColors)
            ? ParseCellColors(JsonNode.Parse(cellColors))
            : [];
        var mergeCellsList = !string.IsNullOrEmpty(mergeCells)
            ? ParseMergeCells(JsonNode.Parse(mergeCells))
            : [];

        for (var i = 0; i < numRows; i++)
        {
            for (var j = 0; j < numCols; j++)
            {
                builder.InsertCell();

                if (builder.CurrentParagraph.ParentNode is Cell cell)
                {
                    cell.CellFormat.VerticalAlignment = GetVerticalAlignment(verticalAlignment);

                    var specificColor = cellColorsList.FirstOrDefault(c => c.row == i && c.col == j);
                    if (!string.IsNullOrEmpty(specificColor.color))
                        cell.CellFormat.Shading.BackgroundPatternColor =
                            ColorHelper.ParseColor(specificColor.color, true);
                    else if (rowColorsDict.TryGetValue(i, out var rowColor))
                        cell.CellFormat.Shading.BackgroundPatternColor = ColorHelper.ParseColor(rowColor, true);
                    else if (hasHeader && i == 0 && !string.IsNullOrEmpty(headerBackgroundColor))
                        cell.CellFormat.Shading.BackgroundPatternColor =
                            ColorHelper.ParseColor(headerBackgroundColor, true);
                    else if (!string.IsNullOrEmpty(alternatingRowColor) && i > 0 && i % 2 == 0)
                        cell.CellFormat.Shading.BackgroundPatternColor =
                            ColorHelper.ParseColor(alternatingRowColor, true);
                    else if (!string.IsNullOrEmpty(cellBackgroundColor))
                        cell.CellFormat.Shading.BackgroundPatternColor =
                            ColorHelper.ParseColor(cellBackgroundColor, true);
                }

                var cellText = "";
                if (parsedTableData != null && i < parsedTableData.Count && j < parsedTableData[i].Count)
                    cellText = parsedTableData[i][j];

                if (fontSize.HasValue)
                    builder.Font.Size = fontSize.Value;

                if (!string.IsNullOrEmpty(fontName))
                    builder.Font.Name = fontName;

                if (hasHeader && i == 0)
                    builder.Font.Bold = true;
                else
                    builder.Font.Bold = false;

                if (!string.IsNullOrEmpty(cellText))
                {
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

            builder.EndRow();
        }

        builder.EndTable();

        if (tableWidth.HasValue)
            table.PreferredWidth = PreferredWidth.FromPoints(tableWidth.Value);

        table.AllowAutoFit = autoFit;

        foreach (var merge in mergeCellsList)
            ApplyMergeCells(table, merge.startRow, merge.endRow, merge.startCol, merge.endCol);

        ctx.Save(outputPath);
        var result = $"Successfully created table with {numRows} rows and {numCols} columns.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Applies merge to cells in specified range.
    /// </summary>
    /// <param name="table">The table to apply merge to.</param>
    /// <param name="startRow">The starting row index.</param>
    /// <param name="endRow">The ending row index.</param>
    /// <param name="startCol">The starting column index.</param>
    /// <param name="endCol">The ending column index.</param>
    private static void ApplyMergeCells(Table table, int startRow, int endRow, int startCol, int endCol)
    {
        if (startRow > endRow || startCol > endCol) return;
        if (startRow < 0 || startRow >= table.Rows.Count) return;
        if (endRow < 0 || endRow >= table.Rows.Count) return;

        for (var row = startRow; row <= endRow; row++)
        {
            var currentRow = table.Rows[row];
            for (var col = startCol; col <= endCol; col++)
            {
                if (col >= currentRow.Cells.Count) continue;

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

    /// <summary>
    ///     Deletes a table from the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="tableIndex">The table index (0-based).</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when tableIndex or sectionIndex is out of range.</exception>
    private static string DeleteTable(DocumentContext<Document> ctx, string? outputPath, int tableIndex,
        int? sectionIndex)
    {
        var doc = ctx.Document;
        List<Table> tables;

        if (sectionIndex.HasValue)
        {
            if (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count)
                throw new ArgumentException($"Section index {sectionIndex.Value} out of range");
            var section = doc.Sections[sectionIndex.Value];
            tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        }
        else
        {
            tables = doc.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        }

        if (tableIndex < 0 || tableIndex >= tables.Count)
            throw new ArgumentException($"Table index {tableIndex} out of range");

        tables[tableIndex].Remove();

        ctx.Save(outputPath);
        var result = $"Successfully deleted table #{tableIndex}. Remaining tables: {tables.Count - 1}.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Gets all tables from the document.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <returns>A JSON string containing table information.</returns>
    /// <exception cref="ArgumentException">Thrown when sectionIndex is out of range.</exception>
    private static string GetTables(DocumentContext<Document> ctx, int? sectionIndex)
    {
        var doc = ctx.Document;
        List<Table> tables;

        if (sectionIndex.HasValue)
        {
            if (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count)
                throw new ArgumentException($"Section index {sectionIndex.Value} out of range");
            var section = doc.Sections[sectionIndex.Value];
            tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        }
        else
        {
            tables = doc.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        }

        List<object> tableList = [];
        for (var i = 0; i < tables.Count; i++)
        {
            var table = tables[i];
            var rowCount = table.Rows.Count;
            var colCount = rowCount > 0 ? table.Rows[0].Cells.Count : 0;

            var precedingText = GetPrecedingText(table, 50);

            tableList.Add(new
            {
                index = i,
                rows = rowCount,
                columns = colCount,
                precedingText = !string.IsNullOrEmpty(precedingText) ? precedingText : null
            });
        }

        var result = new
        {
            count = tables.Count,
            sectionIndex,
            tables = tableList
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Inserts a row into a table.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="tableIndex">The table index (0-based).</param>
    /// <param name="rowIndex">The row index to insert at (0-based).</param>
    /// <param name="rowData">The row data as JSON array.</param>
    /// <param name="insertBefore">Whether to insert before the target position.</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when rowIndex is missing or indices are out of range.</exception>
    private static string InsertRow(DocumentContext<Document> ctx, string? outputPath, int tableIndex, int? rowIndex,
        string? rowData, bool insertBefore, int? sectionIndex)
    {
        if (!rowIndex.HasValue)
            throw new ArgumentException("rowIndex is required for insert_row operation");

        var doc = ctx.Document;
        var actualSectionIndex = sectionIndex ?? 0;
        if (actualSectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"Section index {actualSectionIndex} out of range");

        var section = doc.Sections[actualSectionIndex];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        if (tableIndex < 0 || tableIndex >= tables.Count)
            throw new ArgumentException($"Table index {tableIndex} out of range");

        var table = tables[tableIndex];
        if (rowIndex.Value < 0 || rowIndex.Value >= table.Rows.Count)
            throw new ArgumentException($"Row index {rowIndex.Value} out of range");

        JsonArray? dataArray = null;
        if (!string.IsNullOrEmpty(rowData))
            try
            {
                dataArray = JsonNode.Parse(rowData)?.AsArray();
            }
            catch
            {
                throw new ArgumentException("Invalid rowData JSON format");
            }

        var targetRow = table.Rows[rowIndex.Value];
        var newRow = new Row(doc);

        for (var i = 0; i < targetRow.Cells.Count; i++)
        {
            var sourceCell = targetRow.Cells[i];
            var newCell = new Cell(doc)
            {
                CellFormat =
                {
                    Width = sourceCell.CellFormat.Width,
                    VerticalAlignment = sourceCell.CellFormat.VerticalAlignment
                }
            };
            newCell.CellFormat.SetPaddings(
                sourceCell.CellFormat.TopPadding,
                sourceCell.CellFormat.BottomPadding,
                sourceCell.CellFormat.LeftPadding,
                sourceCell.CellFormat.RightPadding
            );

            newRow.AppendChild(newCell);
            if (dataArray != null && i < dataArray.Count)
            {
                var cellText = dataArray[i]?.GetValue<string>() ?? "";
                if (!string.IsNullOrEmpty(cellText))
                {
                    var para = new Paragraph(doc);
                    if (cellText.Contains('\n'))
                    {
                        var lines = cellText.Split('\n');
                        for (var lineIdx = 0; lineIdx < lines.Length; lineIdx++)
                        {
                            if (lineIdx > 0)
                                para.AppendChild(new Run(doc, ControlChar.LineBreak));
                            para.AppendChild(new Run(doc, lines[lineIdx]));
                        }
                    }
                    else
                    {
                        para.AppendChild(new Run(doc, cellText));
                    }

                    newCell.AppendChild(para);
                }
            }
        }

        if (insertBefore)
            table.InsertBefore(newRow, targetRow);
        else
            table.InsertAfter(newRow, targetRow);

        ctx.Save(outputPath);
        var result = $"Successfully inserted row at index {(insertBefore ? rowIndex.Value : rowIndex.Value + 1)}.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Deletes a row from a table.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="tableIndex">The table index (0-based).</param>
    /// <param name="rowIndex">The row index to delete (0-based).</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when rowIndex is missing or indices are out of range.</exception>
    private static string DeleteRow(DocumentContext<Document> ctx, string? outputPath, int tableIndex, int? rowIndex,
        int? sectionIndex)
    {
        if (!rowIndex.HasValue)
            throw new ArgumentException("rowIndex is required for delete_row operation");

        var doc = ctx.Document;
        var actualSectionIndex = sectionIndex ?? 0;
        if (actualSectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"Section index {actualSectionIndex} out of range");

        var section = doc.Sections[actualSectionIndex];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        if (tableIndex < 0 || tableIndex >= tables.Count)
            throw new ArgumentException($"Table index {tableIndex} out of range");

        var table = tables[tableIndex];
        if (rowIndex.Value < 0 || rowIndex.Value >= table.Rows.Count)
            throw new ArgumentException($"Row index {rowIndex.Value} out of range");

        var rowToDelete = table.Rows[rowIndex.Value];
        rowToDelete.Remove();

        ctx.Save(outputPath);
        var result = $"Successfully deleted row #{rowIndex.Value}. Remaining rows: {table.Rows.Count}.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Inserts a column into a table.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="tableIndex">The table index (0-based).</param>
    /// <param name="columnIndex">The column index to insert at (0-based).</param>
    /// <param name="columnData">The column data as JSON array.</param>
    /// <param name="insertBefore">Whether to insert before the target position.</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when columnIndex is missing or indices are out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when the table has no rows.</exception>
    private static string InsertColumn(DocumentContext<Document> ctx, string? outputPath, int tableIndex,
        int? columnIndex, string? columnData, bool insertBefore, int? sectionIndex)
    {
        if (!columnIndex.HasValue)
            throw new ArgumentException("columnIndex is required for insert_column operation");

        var doc = ctx.Document;
        var actualSectionIndex = sectionIndex ?? 0;
        if (actualSectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"Section index {actualSectionIndex} out of range");

        var section = doc.Sections[actualSectionIndex];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        if (tableIndex < 0 || tableIndex >= tables.Count)
            throw new ArgumentException($"Table index {tableIndex} out of range");

        var table = tables[tableIndex];
        if (table.Rows.Count == 0)
            throw new InvalidOperationException($"Table {tableIndex} has no rows");

        var firstRow = table.Rows[0];
        if (columnIndex.Value < 0 || columnIndex.Value >= firstRow.Cells.Count)
            throw new ArgumentException($"Column index {columnIndex.Value} out of range");

        JsonArray? dataArray = null;
        if (!string.IsNullOrEmpty(columnData))
            try
            {
                dataArray = JsonNode.Parse(columnData)?.AsArray();
            }
            catch
            {
                throw new ArgumentException("Invalid columnData JSON format");
            }

        var insertPosition = insertBefore ? columnIndex.Value : columnIndex.Value + 1;

        for (var rowIdx = 0; rowIdx < table.Rows.Count; rowIdx++)
        {
            var row = table.Rows[rowIdx];
            var newCell = new Cell(doc);
            if (columnIndex.Value < row.Cells.Count)
            {
                var sourceCell = row.Cells[columnIndex.Value];
                newCell.CellFormat.Width = sourceCell.CellFormat.Width;
                newCell.CellFormat.VerticalAlignment = sourceCell.CellFormat.VerticalAlignment;
                newCell.CellFormat.SetPaddings(
                    sourceCell.CellFormat.TopPadding,
                    sourceCell.CellFormat.BottomPadding,
                    sourceCell.CellFormat.LeftPadding,
                    sourceCell.CellFormat.RightPadding
                );
            }

            if (dataArray != null && rowIdx < dataArray.Count)
            {
                var cellDataNode = dataArray[rowIdx];
                var cellText = "";

                if (cellDataNode != null)
                    cellText = cellDataNode.GetValueKind() == JsonValueKind.String
                        ? cellDataNode.GetValue<string>()
                        : cellDataNode.ToString();

                if (!string.IsNullOrEmpty(cellText))
                {
                    var para = new Paragraph(doc);
                    if (cellText.Contains('\n'))
                    {
                        var lines = cellText.Split('\n');
                        for (var lineIdx = 0; lineIdx < lines.Length; lineIdx++)
                        {
                            if (lineIdx > 0)
                                para.AppendChild(new Run(doc, ControlChar.LineBreak));
                            para.AppendChild(new Run(doc, lines[lineIdx]));
                        }
                    }
                    else
                    {
                        para.AppendChild(new Run(doc, cellText));
                    }

                    newCell.AppendChild(para);
                }
            }
            else
            {
                var para = new Paragraph(doc);
                newCell.AppendChild(para);
            }

            if (insertPosition < row.Cells.Count)
            {
                var targetCell = row.Cells[insertPosition];
                row.InsertBefore(newCell, targetCell);
            }
            else
            {
                row.AppendChild(newCell);
            }
        }

        ctx.Save(outputPath);
        var result =
            $"Successfully inserted column at index {(insertBefore ? columnIndex.Value : columnIndex.Value + 1)}.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Deletes a column from a table.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="tableIndex">The table index (0-based).</param>
    /// <param name="columnIndex">The column index to delete (0-based).</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when columnIndex is missing or indices are out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when the table has no rows.</exception>
    private static string DeleteColumn(DocumentContext<Document> ctx, string? outputPath, int tableIndex,
        int? columnIndex, int? sectionIndex)
    {
        if (!columnIndex.HasValue)
            throw new ArgumentException("columnIndex is required for delete_column operation");

        var doc = ctx.Document;
        List<Table> tables;
        if (sectionIndex.HasValue)
        {
            if (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count)
                throw new ArgumentException($"Section index {sectionIndex.Value} out of range");
            var section = doc.Sections[sectionIndex.Value];
            tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        }
        else
        {
            tables = doc.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        }

        if (tableIndex < 0 || tableIndex >= tables.Count)
            throw new ArgumentException($"Table index {tableIndex} out of range");

        var table = tables[tableIndex];
        if (table.Rows.Count == 0)
            throw new InvalidOperationException($"Table {tableIndex} has no rows");

        var firstRow = table.Rows[0];
        if (columnIndex.Value < 0 || columnIndex.Value >= firstRow.Cells.Count)
            throw new ArgumentException($"Column index {columnIndex.Value} out of range");

        var deletedCount = 0;
        foreach (var row in table.Rows.Cast<Row>())
            if (columnIndex.Value < row.Cells.Count)
            {
                row.Cells[columnIndex.Value].Remove();
                deletedCount++;
            }

        ctx.Save(outputPath);
        var result = $"Successfully deleted column #{columnIndex.Value} ({deletedCount} cells removed).\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Merges cells in a table.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="tableIndex">The table index (0-based).</param>
    /// <param name="startRow">The starting row index (0-based).</param>
    /// <param name="startCol">The starting column index (0-based).</param>
    /// <param name="endRow">The ending row index (0-based).</param>
    /// <param name="endCol">The ending column index (0-based).</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or indices are out of range.</exception>
    private static string MergeCells(DocumentContext<Document> ctx, string? outputPath, int tableIndex, int? startRow,
        int? startCol, int? endRow, int? endCol, int? sectionIndex)
    {
        if (!startRow.HasValue || !startCol.HasValue || !endRow.HasValue || !endCol.HasValue)
            throw new ArgumentException(
                "startRow, startCol, endRow, and endCol are all required for merge_cells operation");

        var doc = ctx.Document;
        var actualSectionIndex = sectionIndex ?? 0;
        if (actualSectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"Section index {actualSectionIndex} out of range");

        var section = doc.Sections[actualSectionIndex];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        if (tableIndex < 0 || tableIndex >= tables.Count)
            throw new ArgumentException($"Table index {tableIndex} out of range");

        var table = tables[tableIndex];
        if (startRow.Value < 0 || startRow.Value >= table.Rows.Count || endRow.Value < 0 ||
            endRow.Value >= table.Rows.Count)
            throw new ArgumentException("Row indices out of range");
        if (startRow.Value > endRow.Value)
            throw new ArgumentException($"Start row {startRow.Value} cannot be greater than end row {endRow.Value}");

        var firstRowForCheck = table.Rows[startRow.Value];
        if (startCol.Value < 0 || startCol.Value >= firstRowForCheck.Cells.Count || endCol.Value < 0 ||
            endCol.Value >= firstRowForCheck.Cells.Count)
            throw new ArgumentException("Column indices out of range");
        if (startCol.Value > endCol.Value)
            throw new ArgumentException(
                $"Start column {startCol.Value} cannot be greater than end column {endCol.Value}");

        for (var row = startRow.Value; row <= endRow.Value; row++)
        {
            var currentRow = table.Rows[row];
            for (var col = startCol.Value; col <= endCol.Value; col++)
            {
                var cell = currentRow.Cells[col];
                if (row == startRow.Value && col == startCol.Value)
                {
                    if (startRow.Value != endRow.Value)
                        cell.CellFormat.VerticalMerge = CellMerge.First;
                    if (startCol.Value != endCol.Value)
                        cell.CellFormat.HorizontalMerge = CellMerge.First;
                }
                else
                {
                    if (row == startRow.Value)
                    {
                        cell.CellFormat.HorizontalMerge = CellMerge.Previous;
                    }
                    else if (col == startCol.Value)
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

        ctx.Save(outputPath);
        var result =
            $"Successfully merged cells from [{startRow.Value}, {startCol.Value}] to [{endRow.Value}, {endCol.Value}].\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Splits a cell in a table.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="tableIndex">The table index (0-based).</param>
    /// <param name="rowIndex">The row index of the cell (0-based).</param>
    /// <param name="columnIndex">The column index of the cell (0-based).</param>
    /// <param name="splitRows">The number of rows to split into.</param>
    /// <param name="splitCols">The number of columns to split into.</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or indices are out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when trying to split a merged cell.</exception>
    private static string SplitCell(DocumentContext<Document> ctx, string? outputPath, int tableIndex, int? rowIndex,
        int? columnIndex, int splitRows, int splitCols, int? sectionIndex)
    {
        if (!rowIndex.HasValue)
            throw new ArgumentException("rowIndex is required for split_cell operation");
        if (!columnIndex.HasValue)
            throw new ArgumentException("columnIndex is required for split_cell operation");

        var doc = ctx.Document;
        var actualSectionIndex = sectionIndex ?? 0;
        if (actualSectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"Section index {actualSectionIndex} out of range");

        var section = doc.Sections[actualSectionIndex];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
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

            var para = new Paragraph(doc);
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

        if (splitRows > 1)
            for (var r = 1; r < splitRows; r++)
            {
                var insertAfterRowIndex = rowIndex.Value + r - 1;
                if (insertAfterRowIndex < table.Rows.Count)
                {
                    var newRow = new Row(doc);
                    var sourceRow = table.Rows[rowIndex.Value];
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
                        var paraInRow = new Paragraph(doc);
                        newCellInRow.AppendChild(paraInRow);
                        newRow.AppendChild(newCellInRow);
                    }

                    table.InsertAfter(newRow, table.Rows[insertAfterRowIndex]);
                }
            }

        ctx.Save(outputPath);
        var result =
            $"Successfully split cell [{rowIndex.Value}, {columnIndex.Value}] into {splitRows} rows x {splitCols} columns.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Edits cell format properties.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="tableIndex">The table index (0-based).</param>
    /// <param name="rowIndex">The row index (0-based).</param>
    /// <param name="columnIndex">The column index (0-based).</param>
    /// <param name="applyToRow">Whether to apply formatting to the entire row.</param>
    /// <param name="applyToColumn">Whether to apply formatting to the entire column.</param>
    /// <param name="applyToTable">Whether to apply formatting to the entire table.</param>
    /// <param name="backgroundColor">The background color in hex format.</param>
    /// <param name="alignment">The text alignment.</param>
    /// <param name="verticalAlignmentStr">The vertical alignment.</param>
    /// <param name="paddingTop">The top padding in points.</param>
    /// <param name="paddingBottom">The bottom padding in points.</param>
    /// <param name="paddingLeft">The left padding in points.</param>
    /// <param name="paddingRight">The right padding in points.</param>
    /// <param name="fontName">The font name.</param>
    /// <param name="fontNameAscii">The font name for ASCII characters.</param>
    /// <param name="fontNameFarEast">The font name for Far East characters.</param>
    /// <param name="fontSize">The font size in points.</param>
    /// <param name="bold">Whether the text should be bold.</param>
    /// <param name="italic">Whether the text should be italic.</param>
    /// <param name="color">The text color in hex format.</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or indices are out of range.</exception>
    private static string EditCellFormat(DocumentContext<Document> ctx, string? outputPath, int tableIndex,
        int? rowIndex, int? columnIndex, bool applyToRow, bool applyToColumn, bool applyToTable,
        string? backgroundColor, string? alignment, string? verticalAlignmentStr, double? paddingTop,
        double? paddingBottom, double? paddingLeft, double? paddingRight, string? fontName, string? fontNameAscii,
        string? fontNameFarEast, double? fontSize, bool? bold, bool? italic, string? color, int? sectionIndex)
    {
        var doc = ctx.Document;
        var actualSectionIndex = sectionIndex ?? 0;
        if (actualSectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"Section index {actualSectionIndex} out of range");

        var section = doc.Sections[actualSectionIndex];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        if (tableIndex < 0 || tableIndex >= tables.Count)
            throw new ArgumentException($"Table index {tableIndex} out of range");

        var table = tables[tableIndex];

        List<Cell> targetCells = [];

        if (applyToTable)
        {
            foreach (var row in table.Rows.Cast<Row>())
                targetCells.AddRange(row.Cells.Cast<Cell>());
        }
        else if (applyToRow)
        {
            if (!rowIndex.HasValue)
                throw new ArgumentException("rowIndex is required when applyToRow is true");
            if (rowIndex.Value < 0 || rowIndex.Value >= table.Rows.Count)
                throw new ArgumentException($"Row index {rowIndex.Value} out of range");
            targetCells.AddRange(table.Rows[rowIndex.Value].Cells.Cast<Cell>());
        }
        else if (applyToColumn)
        {
            if (!columnIndex.HasValue)
                throw new ArgumentException("columnIndex is required when applyToColumn is true");
            foreach (var row in table.Rows.Cast<Row>())
                if (columnIndex.Value < row.Cells.Count)
                    targetCells.Add(row.Cells[columnIndex.Value]);
        }
        else
        {
            if (!rowIndex.HasValue)
                throw new ArgumentException("rowIndex is required for edit_cell_format operation");
            if (!columnIndex.HasValue)
                throw new ArgumentException("columnIndex is required for edit_cell_format operation");
            if (rowIndex.Value < 0 || rowIndex.Value >= table.Rows.Count)
                throw new ArgumentException($"Row index {rowIndex.Value} out of range");
            var row = table.Rows[rowIndex.Value];
            if (columnIndex.Value < 0 || columnIndex.Value >= row.Cells.Count)
                throw new ArgumentException($"Column index {columnIndex.Value} out of range");
            targetCells.Add(row.Cells[columnIndex.Value]);
        }

        if (targetCells.Count == 0)
            throw new ArgumentException("No target cells found");

        var hasTextFormatting = !string.IsNullOrEmpty(fontName) || !string.IsNullOrEmpty(fontNameAscii) ||
                                !string.IsNullOrEmpty(fontNameFarEast) || fontSize.HasValue ||
                                bold.HasValue || italic.HasValue || !string.IsNullOrEmpty(color);

        foreach (var cell in targetCells)
        {
            var cellFormat = cell.CellFormat;

            if (!string.IsNullOrEmpty(backgroundColor))
                cellFormat.Shading.BackgroundPatternColor = ColorHelper.ParseColor(backgroundColor, true);

            if (!string.IsNullOrEmpty(alignment))
            {
                var paragraphs = cell.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
                foreach (var para in paragraphs)
                    para.ParagraphFormat.Alignment = alignment.ToLower() switch
                    {
                        "center" => ParagraphAlignment.Center,
                        "right" => ParagraphAlignment.Right,
                        "justify" => ParagraphAlignment.Justify,
                        _ => ParagraphAlignment.Left
                    };
            }

            if (!string.IsNullOrEmpty(verticalAlignmentStr))
                cellFormat.VerticalAlignment = verticalAlignmentStr.ToLower() switch
                {
                    "center" => CellVerticalAlignment.Center,
                    "bottom" => CellVerticalAlignment.Bottom,
                    _ => CellVerticalAlignment.Top
                };

            if (paddingTop.HasValue)
                cellFormat.TopPadding = paddingTop.Value;
            if (paddingBottom.HasValue)
                cellFormat.BottomPadding = paddingBottom.Value;
            if (paddingLeft.HasValue)
                cellFormat.LeftPadding = paddingLeft.Value;
            if (paddingRight.HasValue)
                cellFormat.RightPadding = paddingRight.Value;

            if (hasTextFormatting)
            {
                var runs = cell.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
                foreach (var run in runs)
                    FontHelper.Word.ApplyFontSettings(
                        run,
                        fontName,
                        fontNameAscii,
                        fontNameFarEast,
                        fontSize,
                        bold,
                        italic,
                        color: color
                    );
            }
        }

        var targetDescription = applyToTable
            ? "entire table"
            : applyToRow
                ? $"row {rowIndex}"
                : applyToColumn
                    ? $"column {columnIndex}"
                    : $"cell [{rowIndex}, {columnIndex}]";

        ctx.Save(outputPath);
        var result = $"Successfully edited {targetDescription} format ({targetCells.Count} cells).\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Moves a table to a different position.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="tableIndex">The table index (0-based).</param>
    /// <param name="targetParagraphIndex">The target paragraph index (-1 for end).</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when indices are out of range or target paragraph cannot be found.</exception>
    private static string MoveTable(DocumentContext<Document> ctx, string? outputPath, int tableIndex,
        int targetParagraphIndex, int? sectionIndex)
    {
        var doc = ctx.Document;
        var sectionIdx = sectionIndex ?? 0;
        if (sectionIdx < 0 || sectionIdx >= doc.Sections.Count)
            throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

        var section = doc.Sections[sectionIdx];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, false).Cast<Paragraph>().ToList();

        if (tableIndex < 0 || tableIndex >= tables.Count)
            throw new ArgumentException($"tableIndex must be between 0 and {tables.Count - 1}");

        var table = tables[tableIndex];
        Paragraph? targetPara;

        if (targetParagraphIndex == -1)
        {
            if (paragraphs.Count > 0)
                targetPara = paragraphs[^1];
            else
                throw new ArgumentException(
                    "Cannot move table: section has no paragraphs. Use a valid paragraph index.");
        }
        else if (targetParagraphIndex < 0 || targetParagraphIndex >= paragraphs.Count)
        {
            throw new ArgumentException(
                $"targetParagraphIndex must be between 0 and {paragraphs.Count - 1}, or use -1 for document end");
        }
        else
        {
            targetPara = paragraphs[targetParagraphIndex];
        }

        if (targetPara == null) throw new ArgumentException("Cannot find target paragraph");

        section.Body.InsertAfter(table, targetPara);

        ctx.Save(outputPath);
        var result = $"Successfully moved table {tableIndex}.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Copies a table to another location.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="tableIndex">The table index to copy (0-based).</param>
    /// <param name="targetParagraphIndex">The target paragraph index (-1 for end).</param>
    /// <param name="sourceSectionIndex">The source section index (0-based).</param>
    /// <param name="targetSectionIndex">The target section index (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when indices are out of range or target paragraph cannot be found.</exception>
    private static string CopyTable(DocumentContext<Document> ctx, string? outputPath, int tableIndex,
        int targetParagraphIndex, int? sourceSectionIndex, int? targetSectionIndex)
    {
        var doc = ctx.Document;
        var sourceSectionIdx = sourceSectionIndex ?? 0;
        var targetSectionIdx = targetSectionIndex ?? 0;

        if (sourceSectionIdx < 0 || sourceSectionIdx >= doc.Sections.Count)
            throw new ArgumentException($"sourceSectionIndex must be between 0 and {doc.Sections.Count - 1}");
        if (targetSectionIdx < 0 || targetSectionIdx >= doc.Sections.Count)
            throw new ArgumentException($"targetSectionIndex must be between 0 and {doc.Sections.Count - 1}");

        var sourceSection = doc.Sections[sourceSectionIdx];
        var sourceTables = sourceSection.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        if (tableIndex < 0 || tableIndex >= sourceTables.Count)
            throw new ArgumentException($"sourceTableIndex must be between 0 and {sourceTables.Count - 1}");

        var sourceTable = sourceTables[tableIndex];
        var targetSection = doc.Sections[targetSectionIdx];
        var targetParagraphs =
            targetSection.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

        Paragraph? targetPara;
        if (targetParagraphIndex == -1)
        {
            if (targetParagraphs.Count > 0)
                targetPara = targetParagraphs[^1];
            else
                throw new ArgumentException(
                    "Cannot copy table: target section has no paragraphs. Use a valid paragraph index.");
        }
        else if (targetParagraphIndex < 0 || targetParagraphIndex >= targetParagraphs.Count)
        {
            throw new ArgumentException(
                $"targetParagraphIndex must be between 0 and {targetParagraphs.Count - 1}, or use -1 for document end");
        }
        else
        {
            targetPara = targetParagraphs[targetParagraphIndex];
        }

        if (targetPara == null) throw new ArgumentException("Cannot find target paragraph");

        Node? insertionPoint;

        if (targetPara.ParentNode == targetSection.Body)
        {
            insertionPoint = targetPara;
        }
        else
        {
            var bodyParagraphs = targetSection.Body.GetChildNodes(NodeType.Paragraph, false);
            Paragraph? directPara = null;

            foreach (var para in bodyParagraphs.Cast<Paragraph>())
                if (para == targetPara)
                {
                    directPara = para;
                    break;
                }

            if (directPara == null && bodyParagraphs.Count > 0)
                directPara = bodyParagraphs[^1] as Paragraph;

            insertionPoint = directPara ?? targetSection.Body.LastChild;
        }

        if (insertionPoint == null)
            throw new ArgumentException(
                $"Unable to find valid insertion point (targetParagraphIndex: {targetParagraphIndex})");

        var clonedTable = (Table)sourceTable.Clone(true);
        targetSection.Body.InsertAfter(clonedTable, insertionPoint);

        ctx.Save(outputPath);
        var result = $"Successfully copied table {tableIndex} to paragraph {targetParagraphIndex}.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Gets detailed table structure information.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="tableIndex">The table index (0-based).</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <param name="includeContent">Whether to include content in the output.</param>
    /// <param name="includeCellFormatting">Whether to include cell formatting in the output.</param>
    /// <returns>A formatted string containing table structure information.</returns>
    /// <exception cref="ArgumentException">Thrown when tableIndex or sectionIndex is out of range.</exception>
    private static string GetTableStructure(DocumentContext<Document> ctx, int tableIndex, int? sectionIndex,
        bool includeContent, bool includeCellFormatting)
    {
        var doc = ctx.Document;
        var actualSectionIndex = sectionIndex ?? 0;
        if (actualSectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"Section index {actualSectionIndex} out of range");

        var section = doc.Sections[actualSectionIndex];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        if (tableIndex < 0 || tableIndex >= tables.Count)
            throw new ArgumentException($"Table index {tableIndex} out of range");

        var table = tables[tableIndex];
        var result = new StringBuilder();

        result.AppendLine($"=== Table #{tableIndex} Structure ===\n");
        result.AppendLine("[Basic Info]");
        result.AppendLine($"Rows: {table.Rows.Count}");
        if (table.Rows.Count > 0)
            result.AppendLine($"Columns: {table.Rows[0].Cells.Count}");
        result.AppendLine();

        result.AppendLine("[Table Format]");
        result.AppendLine($"Alignment: {table.Alignment}");
        result.AppendLine($"Style: {table.Style?.Name ?? "None"}");
        result.AppendLine($"Left Indent: {table.LeftIndent:F2} pt");
        if (table.PreferredWidth.Type != PreferredWidthType.Auto)
            result.AppendLine($"Width: {table.PreferredWidth.Value} ({table.PreferredWidth.Type})");
        result.AppendLine($"Allow Auto Fit: {table.AllowAutoFit}");
        result.AppendLine();

        if (includeContent)
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

        if (includeCellFormatting && table.Rows.Count > 0 && table.Rows[0].Cells.Count > 0)
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

        return result.ToString();
    }

    /// <summary>
    ///     Sets table border properties.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="tableIndex">The table index (0-based).</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <param name="rowIndex">The row index (0-based, optional).</param>
    /// <param name="columnIndex">The column index (0-based, optional).</param>
    /// <param name="borderTop">Whether to enable top border.</param>
    /// <param name="borderBottom">Whether to enable bottom border.</param>
    /// <param name="borderLeft">Whether to enable left border.</param>
    /// <param name="borderRight">Whether to enable right border.</param>
    /// <param name="lineStyleStr">The border line style.</param>
    /// <param name="lineWidth">The border line width in points.</param>
    /// <param name="lineColor">The border line color in hex format.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when indices are out of range.</exception>
    private static string SetTableBorder(DocumentContext<Document> ctx, string? outputPath, int tableIndex,
        int? sectionIndex, int? rowIndex, int? columnIndex, bool borderTop, bool borderBottom, bool borderLeft,
        bool borderRight, string lineStyleStr, double lineWidth, string lineColor)
    {
        var doc = ctx.Document;
        var actualSectionIndex = sectionIndex ?? 0;
        if (actualSectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"Section index {actualSectionIndex} out of range");

        var section = doc.Sections[actualSectionIndex];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        if (tableIndex >= tables.Count)
            throw new ArgumentException($"Table index {tableIndex} out of range");

        var table = tables[tableIndex];
        var lineStyleEnum = GetLineStyle(lineStyleStr);
        var lineColorParsed = ColorHelper.ParseColor(lineColor);

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
            foreach (var row in table.Rows.Cast<Row>())
                if (columnIndex.Value < row.Cells.Count)
                    targetCells.Add(row.Cells[columnIndex.Value]);
        }
        else
        {
            foreach (var row in table.Rows.Cast<Row>())
                targetCells.AddRange(row.Cells.Cast<Cell>());
        }

        foreach (var cell in targetCells)
        {
            var borders = cell.CellFormat.Borders;
            if (borderTop)
            {
                borders.Top.LineStyle = lineStyleEnum;
                borders.Top.LineWidth = lineWidth;
                borders.Top.Color = lineColorParsed;
            }

            if (borderBottom)
            {
                borders.Bottom.LineStyle = lineStyleEnum;
                borders.Bottom.LineWidth = lineWidth;
                borders.Bottom.Color = lineColorParsed;
            }

            if (borderLeft)
            {
                borders.Left.LineStyle = lineStyleEnum;
                borders.Left.LineWidth = lineWidth;
                borders.Left.Color = lineColorParsed;
            }

            if (borderRight)
            {
                borders.Right.LineStyle = lineStyleEnum;
                borders.Right.LineWidth = lineWidth;
                borders.Right.Color = lineColorParsed;
            }
        }

        ctx.Save(outputPath);
        var result = $"Successfully set table {tableIndex} borders.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Sets column width for a table.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="tableIndex">The table index (0-based).</param>
    /// <param name="columnIndex">The column index (0-based).</param>
    /// <param name="columnWidth">The column width in points.</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or indices are out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when the table has no rows.</exception>
    private static string SetColumnWidth(DocumentContext<Document> ctx, string? outputPath, int tableIndex,
        int? columnIndex, double? columnWidth, int? sectionIndex)
    {
        if (!columnIndex.HasValue)
            throw new ArgumentException("columnIndex is required for set_column_width operation");
        if (!columnWidth.HasValue)
            throw new ArgumentException("columnWidth is required for set_column_width operation");

        if (columnWidth.Value <= 0)
            throw new ArgumentException($"Column width {columnWidth.Value} must be greater than 0");

        var doc = ctx.Document;
        var actualSectionIndex = sectionIndex ?? 0;
        if (actualSectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"Section index {actualSectionIndex} out of range");

        var section = doc.Sections[actualSectionIndex];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        if (tableIndex < 0 || tableIndex >= tables.Count)
            throw new ArgumentException($"Table index {tableIndex} out of range");

        var table = tables[tableIndex];
        if (table.Rows.Count == 0)
            throw new InvalidOperationException($"Table {tableIndex} has no rows");

        var firstRow = table.Rows[0];
        if (columnIndex.Value < 0 || columnIndex.Value >= firstRow.Cells.Count)
            throw new ArgumentException($"Column index {columnIndex.Value} out of range");

        var cellsUpdated = 0;
        foreach (var row in table.Rows.Cast<Row>())
            if (columnIndex.Value < row.Cells.Count)
            {
                row.Cells[columnIndex.Value].CellFormat.PreferredWidth = PreferredWidth.FromPoints(columnWidth.Value);
                cellsUpdated++;
            }

        ctx.Save(outputPath);
        var result =
            $"Successfully set column {columnIndex.Value} width to {columnWidth.Value} pt ({cellsUpdated} cells updated).\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Sets row height for a table.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="tableIndex">The table index (0-based).</param>
    /// <param name="rowIndex">The row index (0-based).</param>
    /// <param name="rowHeight">The row height in points.</param>
    /// <param name="heightRule">The height rule (auto, atLeast, exactly).</param>
    /// <param name="sectionIndex">The section index (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or indices are out of range.</exception>
    private static string SetRowHeight(DocumentContext<Document> ctx, string? outputPath, int tableIndex, int? rowIndex,
        double? rowHeight, string heightRule, int? sectionIndex)
    {
        if (!rowIndex.HasValue)
            throw new ArgumentException("rowIndex is required for set_row_height operation");
        if (!rowHeight.HasValue)
            throw new ArgumentException("rowHeight is required for set_row_height operation");

        if (rowHeight.Value <= 0)
            throw new ArgumentException($"Row height {rowHeight.Value} must be greater than 0");

        var doc = ctx.Document;
        var actualSectionIndex = sectionIndex ?? 0;
        if (actualSectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"Section index {actualSectionIndex} out of range");

        var section = doc.Sections[actualSectionIndex];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
        if (tableIndex < 0 || tableIndex >= tables.Count)
            throw new ArgumentException($"Table index {tableIndex} out of range");

        var table = tables[tableIndex];
        if (rowIndex.Value < 0 || rowIndex.Value >= table.Rows.Count)
            throw new ArgumentException($"Row index {rowIndex.Value} out of range");

        var row = table.Rows[rowIndex.Value];
        row.RowFormat.HeightRule = heightRule.ToLower() switch
        {
            "auto" => HeightRule.Auto,
            "atleast" => HeightRule.AtLeast,
            "exactly" => HeightRule.Exactly,
            _ => HeightRule.AtLeast
        };
        row.RowFormat.Height = rowHeight.Value;

        ctx.Save(outputPath);
        var result = $"Successfully set row {rowIndex.Value} height to {rowHeight.Value} pt ({heightRule}).\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Parses a JSON node into a dictionary of row index to color string
    /// </summary>
    /// <param name="node">JSON node containing row color mappings</param>
    /// <returns>Dictionary mapping row indices to color strings</returns>
    private static Dictionary<int, string> ParseColorDictionary(JsonNode? node)
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
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[WARN] Error parsing cell colors JSON: {ex.Message}");
        }

        return result;
    }

    /// <summary>
    ///     Parses a JSON node into a list of cell color specifications
    /// </summary>
    /// <param name="node">JSON node containing cell color array</param>
    /// <returns>List of tuples containing row, column, and color string</returns>
    private static List<(int row, int col, string color)> ParseCellColors(JsonNode? node)
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
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[WARN] Error parsing merge cells JSON: {ex.Message}");
        }

        return result;
    }

    /// <summary>
    ///     Parses a JSON node into a list of cell merge specifications
    /// </summary>
    /// <param name="node">JSON node containing merge cell specifications</param>
    /// <returns>List of tuples containing start/end row and column indices</returns>
    private static List<(int startRow, int endRow, int startCol, int endCol)> ParseMergeCells(JsonNode? node)
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
        catch (Exception ex)
        {
            Console.Error.WriteLine($"[WARN] Error parsing merge cells JSON: {ex.Message}");
        }

        return result;
    }

    /// <summary>
    ///     Converts a vertical alignment string to CellVerticalAlignment enum
    /// </summary>
    /// <param name="alignment">Vertical alignment string (top, center, bottom)</param>
    /// <returns>Corresponding CellVerticalAlignment value</returns>
    private static CellVerticalAlignment GetVerticalAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "top" => CellVerticalAlignment.Top,
            "bottom" => CellVerticalAlignment.Bottom,
            _ => CellVerticalAlignment.Center
        };
    }

    /// <summary>
    ///     Converts a line style string to LineStyle enum
    /// </summary>
    /// <param name="style">Line style string (none, single, double, dotted, dashed, thick)</param>
    /// <returns>Corresponding LineStyle value</returns>
    private static LineStyle GetLineStyle(string style)
    {
        return style.ToLower() switch
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
    ///     Gets the text content preceding a table node for context identification.
    /// </summary>
    /// <param name="table">The table to get preceding text for.</param>
    /// <param name="maxLength">The maximum length of text to return.</param>
    /// <returns>The preceding text content, truncated if necessary.</returns>
    private static string GetPrecedingText(Table table, int maxLength)
    {
        var precedingSibling = table.PreviousSibling;
        while (precedingSibling != null)
        {
            if (precedingSibling is Paragraph para)
            {
                var text = para.GetText().Trim();
                if (!string.IsNullOrWhiteSpace(text) && !text.StartsWith("\f"))
                {
                    text = text.Replace("\r", "").Replace("\a", "");
                    if (text.Length > maxLength)
                        return text.Substring(0, maxLength) + "...";
                    return text;
                }
            }

            precedingSibling = precedingSibling.PreviousSibling;
        }

        return string.Empty;
    }
}