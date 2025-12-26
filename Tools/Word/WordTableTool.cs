using System.Drawing;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Tables;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Word;

public class WordTableTool : IAsposeTool
{
    public string Description =>
        @"Manage tables in Word documents. Supports 17 operations: add_table, edit_table_format, delete_table, get_tables, insert_row, delete_row, insert_column, delete_column, merge_cells, split_cell, edit_cell_format, move_table, copy_table, get_table_structure, set_table_border, set_column_width, set_row_height.

Usage examples:
- Add table: word_table(operation='add_table', path='doc.docx', rows=3, columns=3, data=[['A1','B1','C1'],['A2','B2','C2']])
- Get tables: word_table(operation='get_tables', path='doc.docx')
- Insert row: word_table(operation='insert_row', path='doc.docx', tableIndex=0, rowIndex=1)
- Merge cells: word_table(operation='merge_cells', path='doc.docx', tableIndex=0, startRow=0, startColumn=0, endRow=0, endColumn=1)
- Set border: word_table(operation='set_table_border', path='doc.docx', tableIndex=0, borderType='all', style='single', width=1.0)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add_table': Add a new table (required params: path, rows, columns)
- 'edit_table_format': Edit table format (required params: path, tableIndex)
- 'delete_table': Delete a table (required params: path, tableIndex)
- 'get_tables': Get all tables info (required params: path)
- 'insert_row': Insert a row (required params: path, tableIndex, rowIndex)
- 'delete_row': Delete a row (required params: path, tableIndex, rowIndex)
- 'insert_column': Insert a column (required params: path, tableIndex, columnIndex)
- 'delete_column': Delete a column (required params: path, tableIndex, columnIndex)
- 'merge_cells': Merge cells (required params: path, tableIndex, startRow, startColumn, endRow, endColumn)
- 'split_cell': Split a cell (required params: path, tableIndex, rowIndex, columnIndex)
- 'edit_cell_format': Edit cell format (required params: path, tableIndex, rowIndex, columnIndex)
- 'move_table': Move table position (required params: path, tableIndex)
- 'copy_table': Copy table (required params: path, tableIndex)
- 'get_table_structure': Get table structure (required params: path, tableIndex)
- 'set_table_border': Set table border (required params: path, tableIndex)
- 'set_column_width': Set column width (required params: path, tableIndex, columnIndex, width)
- 'set_row_height': Set row height (required params: path, tableIndex, rowIndex, height)",
                @enum = new[]
                {
                    "add_table", "edit_table_format", "delete_table", "get_tables", "insert_row", "delete_row",
                    "insert_column", "delete_column", "merge_cells", "split_cell", "edit_cell_format", "move_table",
                    "copy_table", "get_table_structure", "set_table_border", "set_column_width", "set_row_height"
                }
            },
            path = new
            {
                type = "string",
                description = "Document file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to overwrite input)"
            },
            tableIndex = new
            {
                type = "number",
                description =
                    "Table index (0-based, for most operations). Also accepts 'sourceTableIndex' for copy_table operation. Note: After delete operations, subsequent table indices will shift automatically. Use 'get_tables' operation to refresh indices."
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based, default: 0)"
            },
            rows = new
            {
                type = "number",
                description = "Number of rows (for add_table)"
            },
            columns = new
            {
                type = "number",
                description = "Number of columns (for add_table)"
            },
            data = new
            {
                type = "array",
                description = "Table data (array of arrays, for add_table)",
                items = new
                {
                    type = "array",
                    items = new { type = "string" }
                }
            },
            headerRow = new
            {
                type = "boolean",
                description = "First row is header (for add_table, default: false)"
            },
            headerBackgroundColor = new
            {
                type = "string",
                description = "Header row background color hex (for add_table)"
            },
            rowBackgroundColors = new
            {
                type = "object",
                description = "Background colors for specific rows (for add_table)"
            },
            columnBackgroundColors = new
            {
                type = "object",
                description = "Background colors for specific columns (for add_table)"
            },
            cellBackgroundColors = new
            {
                type = "array",
                description = "Background colors for specific cells [row, col, color] (for add_table)",
                items = new
                {
                    type = "array",
                    items = new { type = "string" }
                }
            },
            mergeCells = new
            {
                type = "array",
                description = "Cells to merge (for add_table)",
                items = new
                {
                    type = "object",
                    properties = new
                    {
                        startRow = new { type = "number" },
                        endRow = new { type = "number" },
                        startCol = new { type = "number" },
                        endCol = new { type = "number" }
                    }
                }
            },
            borderStyle = new
            {
                type = "string",
                description = "Border style: none, single, double, dotted (for add_table, default: single)",
                @enum = new[] { "none", "single", "double", "dotted" }
            },
            alignment = new
            {
                type = "string",
                description = @"Table alignment: left, center, right (for add_table, edit_table_format, optional).
Behavior is consistent for both operations:
  - If alignment is provided, it will override the style's default alignment (if styleName is also provided).
  - If alignment is not provided and styleName is provided, the style's default alignment will be used.
  - If neither alignment nor styleName is provided (add_table only), default is 'left'.
  - User-specified parameters always take precedence over style defaults.",
                @enum = new[] { "left", "center", "right" }
            },
            verticalAlignment = new
            {
                type = "string",
                description = @"Cell vertical alignment: top, center, bottom (for add_table, default: center).
Applied to all cells in the table. If styleName is also provided, verticalAlignment will be applied after the style to ensure it takes effect.",
                @enum = new[] { "top", "center", "bottom" }
            },
            cellPadding = new
            {
                type = "number",
                description =
                    "Cell padding in points (for add_table, default: 5, used if individual padding values not specified)"
            },
            tableFontName = new
            {
                type = "string",
                description = "Font name for all table cells (for add_table)"
            },
            tableFontSize = new
            {
                type = "number",
                description = "Font size for all table cells in points (for add_table)"
            },
            tableFontNameAscii = new
            {
                type = "string",
                description = "Font name for English text in table cells (for add_table)"
            },
            tableFontNameFarEast = new
            {
                type = "string",
                description = "Font name for Chinese/Japanese/Korean text in table cells (for add_table)"
            },
            width = new
            {
                type = "number",
                description = "Table width in points (for edit_table_format)"
            },
            widthType = new
            {
                type = "string",
                description = "Width type: auto, points, percent (for edit_table_format)",
                @enum = new[] { "auto", "points", "percent" }
            },
            styleName = new
            {
                type = "string",
                description = @"Table style name (for add_table, edit_table_format).
Behavior is consistent for both operations:
  - Style is applied first, then other format parameters (alignment, verticalAlignment, etc.) are applied to override style defaults.
  - User-specified parameters always take precedence over style defaults.
  - This ensures predictable behavior: explicitly provided parameters will always be respected."
            },
            includeContent = new
            {
                type = "boolean",
                description = "Include table content (for get_tables, get_table_structure, default: false)"
            },
            includeCellFormatting = new
            {
                type = "boolean",
                description = "Include cell-level formatting details (for get_table_structure, default: true)"
            },
            rowIndex = new
            {
                type = "number",
                description =
                    "Row index (0-based, for insert_row, delete_row, merge_cells, split_cell, edit_cell_format, set_row_height)"
            },
            colIndex = new
            {
                type = "number",
                description =
                    "Column index (0-based, for insert_column, delete_column, merge_cells, split_cell, edit_cell_format, set_column_width)"
            },
            insertBefore = new
            {
                type = "boolean",
                description = "Insert before specified index (for insert_row, insert_column, default: false)"
            },
            rowData = new
            {
                type = "array",
                description = "Array of cell data for new row (for insert_row)",
                items = new { type = "string" }
            },
            columnData = new
            {
                type = "array",
                description = "Array of cell data for new column (for insert_column)",
                items = new { type = "string" }
            },
            startRow = new
            {
                type = "number",
                description = "Start row index for merge (for merge_cells)"
            },
            startCol = new
            {
                type = "number",
                description = "Start column index for merge (for merge_cells)"
            },
            endRow = new
            {
                type = "number",
                description = "End row index for merge (for merge_cells)"
            },
            endCol = new
            {
                type = "number",
                description = "End column index for merge (for merge_cells)"
            },
            splitRows = new
            {
                type = "number",
                description = "Number of rows to split into (for split_cell, default: 2)"
            },
            splitCols = new
            {
                type = "number",
                description = "Number of columns to split into (for split_cell, default: 2)"
            },
            backgroundColor = new
            {
                type = "string",
                description = "Cell background color hex (for edit_cell_format)"
            },
            paddingTop = new
            {
                type = "number",
                description = @"Top padding in points (for add_table: default 0.00, for edit_cell_format: optional).
For add_table: Applied to all cells. If individual padding values are not specified, cellPadding will be used for all sides."
            },
            paddingBottom = new
            {
                type = "number",
                description = @"Bottom padding in points (for add_table: default 0.00, for edit_cell_format: optional).
For add_table: Applied to all cells. If individual padding values are not specified, cellPadding will be used for all sides."
            },
            paddingLeft = new
            {
                type = "number",
                description = @"Left padding in points (for add_table: default 5.40, for edit_cell_format: optional).
For add_table: Applied to all cells. If individual padding values are not specified, cellPadding will be used for all sides."
            },
            paddingRight = new
            {
                type = "number",
                description = @"Right padding in points (for add_table: default 5.40, for edit_cell_format: optional).
For add_table: Applied to all cells. If individual padding values are not specified, cellPadding will be used for all sides."
            },
            fontName = new
            {
                type = "string",
                description = "Font name (for edit_cell_format)"
            },
            fontNameAscii = new
            {
                type = "string",
                description = "Font name for ASCII characters (for edit_cell_format)"
            },
            fontNameFarEast = new
            {
                type = "string",
                description = "Font name for Far East characters (for edit_cell_format)"
            },
            fontSize = new
            {
                type = "number",
                description = "Font size in points (for edit_cell_format)"
            },
            bold = new
            {
                type = "boolean",
                description = "Bold text (for edit_cell_format)"
            },
            italic = new
            {
                type = "boolean",
                description = "Italic text (for edit_cell_format)"
            },
            color = new
            {
                type = "string",
                description = "Text color hex (for edit_cell_format)"
            },
            targetParagraphIndex = new
            {
                type = "number",
                description = "Target paragraph index to move/copy after (for move_table, copy_table)"
            },
            sourceTableIndex = new
            {
                type = "number",
                description = "Source table index (for copy_table)"
            },
            sourceSectionIndex = new
            {
                type = "number",
                description = "Source section index (for copy_table)"
            },
            targetSectionIndex = new
            {
                type = "number",
                description = "Target section index (for copy_table)"
            },
            borderTop = new
            {
                type = "boolean",
                description = "Show top border (for set_table_border, default: false)"
            },
            borderBottom = new
            {
                type = "boolean",
                description = "Show bottom border (for set_table_border, default: false)"
            },
            borderLeft = new
            {
                type = "boolean",
                description = "Show left border (for set_table_border, default: false)"
            },
            borderRight = new
            {
                type = "boolean",
                description = "Show right border (for set_table_border, default: false)"
            },
            lineStyle = new
            {
                type = "string",
                description = "Border line style: none, single, double, dotted, dashed, thick (for set_table_border)",
                @enum = new[] { "none", "single", "double", "dotted", "dashed", "thick" }
            },
            lineWidth = new
            {
                type = "number",
                description = "Border line width in points (for set_table_border, default: 0.5)"
            },
            lineColor = new
            {
                type = "string",
                description = "Border line color hex (for set_table_border, default: 000000)"
            },
            columnWidth = new
            {
                type = "number",
                description = "Column width in points (for set_column_width)"
            },
            rowHeight = new
            {
                type = "number",
                description = "Row height in points (for set_row_height)"
            },
            heightRule = new
            {
                type = "string",
                description = "Height rule: auto, atLeast, exactly (for set_row_height, default: atLeast)",
                @enum = new[] { "auto", "atLeast", "exactly" }
            },
            allowAutoFit = new
            {
                type = "boolean",
                description = @"Allow auto fit for table (for add_table, default: true).
When true, the table will automatically adjust column widths to fit content. When false, column widths remain fixed."
            },
            applyToRow = new
            {
                type = "boolean",
                description = @"Apply format to entire row (for edit_cell_format, default: false).
When true, format settings will be applied to all cells in the specified row. rowIndex is required when this is true."
            },
            applyToColumn = new
            {
                type = "boolean",
                description = @"Apply format to entire column (for edit_cell_format, default: false).
When true, format settings will be applied to all cells in the specified column. colIndex is required when this is true."
            },
            applyToTable = new
            {
                type = "boolean",
                description = @"Apply format to entire table (for edit_cell_format, default: false).
When true, format settings will be applied to all cells in the table. This is useful for bulk formatting operations."
            },
            cellStyles = new
            {
                type = "array",
                description =
                    "Cell paragraph styles (for add_table, 2D array matching data structure, each element: {styleName: 'Normal'})",
                items = new
                {
                    type = "array",
                    items = new
                    {
                        type = "object",
                        properties = new
                        {
                            styleName = new
                                { type = "string", description = "Paragraph style name (e.g., 'Normal', '!序號')" }
                        }
                    }
                }
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetStringNullable(arguments, "outputPath") ?? path;
        SecurityHelper.ValidateFilePath(outputPath, "outputPath", true);

        return operation.ToLower() switch
        {
            "add_table" => await AddTable(path, outputPath, arguments),
            "edit_table_format" => await EditTableFormat(path, outputPath, arguments),
            "delete_table" => await DeleteTable(path, outputPath, arguments),
            "get_tables" => await GetTables(path, arguments),
            "insert_row" => await InsertRow(path, outputPath, arguments),
            "delete_row" => await DeleteRow(path, outputPath, arguments),
            "insert_column" => await InsertColumn(path, outputPath, arguments),
            "delete_column" => await DeleteColumn(path, outputPath, arguments),
            "merge_cells" => await MergeCells(path, outputPath, arguments),
            "split_cell" => await SplitCell(path, outputPath, arguments),
            "edit_cell_format" => await EditCellFormat(path, outputPath, arguments),
            "move_table" => await MoveTable(path, outputPath, arguments),
            "copy_table" => await CopyTable(path, outputPath, arguments),
            "get_table_structure" => await GetTableStructure(path, arguments),
            "set_table_border" => await SetTableBorder(path, outputPath, arguments),
            "set_column_width" => await SetColumnWidth(path, outputPath, arguments),
            "set_row_height" => await SetRowHeight(path, outputPath, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a new table to the document
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing rows, columns, optional data, headerRow, formatting options</param>
    /// <returns>Success message with table index</returns>
    private Task<string> AddTable(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var rows = ArgumentHelper.GetInt(arguments, "rows");
            var columns = ArgumentHelper.GetInt(arguments, "columns");
            var headerRow = ArgumentHelper.GetBool(arguments, "headerRow", false);
            var headerBgColor = ArgumentHelper.GetStringNullable(arguments, "headerBackgroundColor");
            var borderStyle = ArgumentHelper.GetString(arguments, "borderStyle", "single");
            // Alignment: only set if explicitly provided, otherwise let style default take effect
            var alignment = ArgumentHelper.GetStringNullable(arguments, "alignment");
            // Vertical alignment: default to "center" if not specified
            var verticalAlignment = ArgumentHelper.GetString(arguments, "verticalAlignment", "center");

            // Cell padding: support individual padding values or unified cellPadding
            var paddingTop = ArgumentHelper.GetDoubleNullable(arguments, "paddingTop");
            var paddingBottom = ArgumentHelper.GetDoubleNullable(arguments, "paddingBottom");
            var paddingLeft = ArgumentHelper.GetDoubleNullable(arguments, "paddingLeft");
            var paddingRight = ArgumentHelper.GetDoubleNullable(arguments, "paddingRight");
            var cellPadding = ArgumentHelper.GetDoubleNullable(arguments, "cellPadding");

            if (!paddingTop.HasValue && !paddingBottom.HasValue && !paddingLeft.HasValue && !paddingRight.HasValue)
            {
                var unifiedPadding = cellPadding ?? 5.0;
                paddingTop = unifiedPadding;
                paddingBottom = unifiedPadding;
                paddingLeft = unifiedPadding;
                paddingRight = unifiedPadding;
            }
            else
            {
                paddingTop ??= 0.0;
                paddingBottom ??= 0.0;
                paddingLeft ??= 5.4;
                paddingRight ??= 5.4;
            }

            var tableFontName = ArgumentHelper.GetStringNullable(arguments, "tableFontName");
            var tableFontSize = ArgumentHelper.GetDoubleNullable(arguments, "tableFontSize");
            var tableFontNameAscii = ArgumentHelper.GetStringNullable(arguments, "tableFontNameAscii");
            var tableFontNameFarEast = ArgumentHelper.GetStringNullable(arguments, "tableFontNameFarEast");
            var allowAutoFit = ArgumentHelper.GetBool(arguments, "allowAutoFit", true);
            var styleName = ArgumentHelper.GetStringNullable(arguments, "styleName");

            var doc = new Document(path);
            var builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();
            builder.InsertParagraph();
            builder.CurrentParagraph.ParagraphFormat.LeftIndent = 0;
            builder.CurrentParagraph.ParagraphFormat.RightIndent = 0;
            builder.CurrentParagraph.ParagraphFormat.FirstLineIndent = 0;

            string[][]? data = null;
            if (arguments?.ContainsKey("data") == true)
                try
                {
                    var dataJson = arguments["data"]?.ToJsonString();
                    if (!string.IsNullOrEmpty(dataJson))
                        data = JsonSerializer.Deserialize<string[][]>(dataJson);
                }
                catch (Exception jsonEx)
                {
                    throw new ArgumentException(
                        $"Unable to parse data parameter: {jsonEx.Message}. Please ensure data is a valid 2D string array format, e.g.: [[\"A1\",\"B1\"],[\"A2\",\"B2\"]]");
                }

            var rowBgColors = ParseColorDictionary(arguments?["rowBackgroundColors"]);
            var columnBgColors = ParseColorDictionary(arguments?["columnBackgroundColors"]);
            var cellColors = ParseCellColors(arguments?["cellBackgroundColors"]);
            var mergeCells = ParseMergeCells(arguments?["mergeCells"]);

            // Parse cell styles (2D array matching data structure)
            Dictionary<(int row, int col), string>? cellStyles = null;
            if (arguments?.ContainsKey("cellStyles") == true)
                try
                {
                    var stylesJson = arguments["cellStyles"]?.ToJsonString();
                    if (!string.IsNullOrEmpty(stylesJson))
                    {
                        var stylesArray = JsonSerializer.Deserialize<JsonElement[][]>(stylesJson);
                        if (stylesArray != null)
                        {
                            cellStyles = new Dictionary<(int row, int col), string>();
                            for (var i = 0; i < stylesArray.Length; i++)
                            {
                                var rowStyles = stylesArray[i];
                                // rowStyles is JsonElement[] which cannot be null according to nullable reference types
                                for (var j = 0; j < rowStyles.Length; j++)
                                    if (rowStyles[j].ValueKind == JsonValueKind.Object &&
                                        rowStyles[j].TryGetProperty("styleName", out var styleNameProp))
                                    {
                                        var cellStyleName = styleNameProp.GetString();
                                        if (!string.IsNullOrEmpty(cellStyleName))
                                            cellStyles[(i, j)] = cellStyleName;
                                    }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new ArgumentException($"Unable to parse cellStyles parameter: {ex.Message}");
                }

            var table = builder.StartTable();
            var cells = new Dictionary<(int row, int col), Cell>();

            for (var i = 0; i < rows; i++)
            {
                for (var j = 0; j < columns; j++)
                {
                    var cell = builder.InsertCell();
                    cells[(i, j)] = cell;

                    var cellText = data != null && i < data.Length && j < data[i].Length
                        ? data[i][j]
                        : $"Cell {i + 1},{j + 1}";

                    // Apply font settings using FontHelper (same logic for header and regular rows)
                    FontHelper.Word.ApplyFontSettings(
                        builder,
                        tableFontName,
                        tableFontNameAscii,
                        tableFontNameFarEast,
                        tableFontSize
                    );

                    // Handle multi-line text: split by \n and insert line breaks
                    if (!string.IsNullOrEmpty(cellText) && cellText.Contains('\n'))
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

                    builder.Font.Bold = false;
                    builder.Font.Italic = false;
                    builder.Font.Color = Color.Black;
                    builder.Font.Name = "Calibri";
                    builder.Font.Size = 11;

                    // SetPaddings parameter order: left, top, right, bottom (Aspose.Words API signature)
                    cell.CellFormat.SetPaddings(paddingLeft.Value, paddingTop.Value, paddingRight.Value,
                        paddingBottom.Value);
                    cell.CellFormat.Shading.BackgroundPatternColor = Color.Empty;
                    cell.CellFormat.VerticalAlignment = GetVerticalAlignment(verticalAlignment);

                    var hasColor = false;
                    var cellColorMatch = cellColors.FirstOrDefault(c => c.row == i && c.col == j);
                    if (cellColorMatch != default)
                    {
                        cell.CellFormat.Shading.BackgroundPatternColor = ColorHelper.ParseColor(cellColorMatch.color);
                        hasColor = true;
                    }

                    if (!hasColor && columnBgColors.TryGetValue(j, out var columnColor))
                    {
                        cell.CellFormat.Shading.BackgroundPatternColor = ColorHelper.ParseColor(columnColor);
                        hasColor = true;
                    }

                    if (!hasColor && rowBgColors.TryGetValue(i, out var rowColor))
                    {
                        cell.CellFormat.Shading.BackgroundPatternColor = ColorHelper.ParseColor(rowColor);
                        hasColor = true;
                    }

                    if (!hasColor && headerRow && i == 0 && !string.IsNullOrEmpty(headerBgColor))
                        cell.CellFormat.Shading.BackgroundPatternColor = ColorHelper.ParseColor(headerBgColor);

                    if (borderStyle != "none")
                    {
                        var lineStyle = borderStyle switch
                        {
                            "double" => LineStyle.Double,
                            "dotted" => LineStyle.Dot,
                            _ => LineStyle.Single
                        };
                        cell.CellFormat.Borders.LineStyle = lineStyle;
                        cell.CellFormat.Borders.Color = Color.Black;
                    }
                    else
                    {
                        cell.CellFormat.Borders.LineStyle = LineStyle.None;
                    }

                    if (cellStyles != null && cellStyles.TryGetValue((i, j), out var cellStyleName))
                        try
                        {
                            var paragraphs = cell.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
                            foreach (var para in paragraphs) para.ParagraphFormat.Style = doc.Styles[cellStyleName];
                        }
                        catch (Exception styleEx)
                        {
                            // Log warning but don't fail - style might not exist
                            Console.Error.WriteLine(
                                $"[WARN] Unable to apply style '{cellStyleName}' to cell [{i}, {j}]: {styleEx.Message}");
                        }
                }

                builder.EndRow();
            }

            builder.EndTable();

            foreach (var merge in mergeCells)
                try
                {
                    var startCell = cells[(merge.startRow, merge.startCol)];
                    if (merge.startRow == merge.endRow && merge.startCol != merge.endCol)
                    {
                        startCell.CellFormat.HorizontalMerge = CellMerge.First;
                        for (var col = merge.startCol + 1; col <= merge.endCol; col++)
                            if (cells.TryGetValue((merge.startRow, col), out var horizontalCell))
                                horizontalCell.CellFormat.HorizontalMerge = CellMerge.Previous;
                    }
                    else if (merge.startCol == merge.endCol && merge.startRow != merge.endRow)
                    {
                        startCell.CellFormat.VerticalMerge = CellMerge.First;
                        for (var row = merge.startRow + 1; row <= merge.endRow; row++)
                            if (cells.TryGetValue((row, merge.startCol), out var verticalCell))
                                verticalCell.CellFormat.VerticalMerge = CellMerge.Previous;
                    }
                }
                catch (Exception ex)
                {
                    // Merge operation failed, but continue with table creation
                    Console.Error.WriteLine($"[WARN] Merge operation failed during table creation: {ex.Message}");
                }

            // Note: In add_table, style is applied first, then other parameters override style defaults
            // This ensures user-specified parameters (alignment, verticalAlignment) take precedence
            if (!string.IsNullOrEmpty(styleName))
                try
                {
                    table.Style = doc.Styles[styleName];
                }
                catch (Exception styleEx)
                {
                    throw new ArgumentException(
                        $"Unable to apply table style '{styleName}': {styleEx.Message}. Use word_get_styles tool to view available styles",
                        styleEx);
                }

            // Set table alignment (only if explicitly provided, overrides style default if style was applied)
            if (!string.IsNullOrEmpty(alignment))
            {
                table.Alignment = alignment.ToLower() switch
                {
                    "center" => TableAlignment.Center,
                    "right" => TableAlignment.Right,
                    _ => TableAlignment.Left
                };
            }
            else
            {
                if (string.IsNullOrEmpty(styleName)) table.Alignment = TableAlignment.Left;
            }

            // Ensure all cells have correct vertical alignment
            // Re-apply after style to ensure it's not overridden by style defaults
            // This is necessary because table styles may include default vertical alignment
            var targetVerticalAlignment = GetVerticalAlignment(verticalAlignment);
            foreach (var row in table.Rows.Cast<Row>())
            foreach (var cell in row.Cells.Cast<Cell>())
                cell.CellFormat.VerticalAlignment = targetVerticalAlignment;

            table.AllowAutoFit = allowAutoFit;

            doc.Save(outputPath);
            return $"Successfully added table ({rows} rows x {columns} columns). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits table format properties
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing tableIndex, optional formatting options</param>
    /// <returns>Success message</returns>
    private Task<string> EditTableFormat(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var tableIndex = ArgumentHelper.GetInt(arguments, "tableIndex");
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);

            var doc = new Document(path);
            if (sectionIndex >= doc.Sections.Count)
                throw new ArgumentException($"Section index {sectionIndex} out of range");

            var section = doc.Sections[sectionIndex];
            var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
            if (tableIndex >= tables.Count)
                throw new ArgumentException($"Table index {tableIndex} out of range");

            var table = tables[tableIndex];

            // Note: In edit_table_format, alignment takes precedence over style default if both are provided
            // This ensures consistency with add_table behavior: user-specified parameters override style defaults

            var styleName = ArgumentHelper.GetStringNullable(arguments, "styleName");
            if (!string.IsNullOrEmpty(styleName))
                try
                {
                    table.Style = doc.Styles[styleName];
                }
                catch (Exception styleEx)
                {
                    throw new ArgumentException(
                        $"Unable to apply table style '{styleName}': {styleEx.Message}. Use word_get_styles tool to view available styles");
                }

            if (arguments?["width"] != null && ArgumentHelper.GetStringNullable(arguments, "widthType") == "points")
            {
                var width = ArgumentHelper.GetDoubleNullable(arguments, "width");
                if (width.HasValue)
                    table.PreferredWidth = PreferredWidth.FromPoints(width.Value);
            }
            else if (ArgumentHelper.GetStringNullable(arguments, "widthType") == "auto")
            {
                table.PreferredWidth = PreferredWidth.Auto;
            }

            // Set alignment last (overrides style default if style was applied)
            var alignment = ArgumentHelper.GetStringNullable(arguments, "alignment");
            if (!string.IsNullOrEmpty(alignment))
                table.Alignment = alignment.ToLower() switch
                {
                    "center" => TableAlignment.Center,
                    "right" => TableAlignment.Right,
                    _ => TableAlignment.Left
                };

            doc.Save(outputPath);
            return $"Successfully edited table {tableIndex} format. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes a table from the document
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing tableIndex, optional sectionIndex</param>
    /// <returns>Success message with remaining table count</returns>
    private Task<string> DeleteTable(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var tableIndex = ArgumentHelper.GetInt(arguments, "tableIndex");
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);

            var doc = new Document(path);
            if (sectionIndex >= doc.Sections.Count)
                throw new ArgumentException($"Section index {sectionIndex} out of range");

            var section = doc.Sections[sectionIndex];
            var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
            if (tableIndex >= tables.Count)
                throw new ArgumentException($"Table index {tableIndex} out of range");

            var table = tables[tableIndex];
            var rowCount = table.Rows.Count;
            var colCount = table.Rows.Count > 0 ? table.Rows[0].Cells.Count : 0;
            table.Remove();

            doc.Save(outputPath);
            return
                $"Successfully deleted table #{tableIndex} ({rowCount} rows x {colCount} columns). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all tables from the document
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="arguments">JSON arguments containing optional sectionIndex, includeContent</param>
    /// <returns>Formatted string with all tables</returns>
    private Task<string> GetTables(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var sectionIndex = ArgumentHelper.GetIntNullable(arguments, "sectionIndex");
            var includeContent = ArgumentHelper.GetBool(arguments, "includeContent", false);

            var doc = new Document(path);
            List<Table> tables;
            if (sectionIndex.HasValue)
            {
                if (sectionIndex.Value < 0 || sectionIndex.Value >= doc.Sections.Count)
                    throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");
                tables = doc.Sections[sectionIndex.Value].Body.GetChildNodes(NodeType.Table, true).Cast<Table>()
                    .ToList();
            }
            else
            {
                tables = doc.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
            }

            var sb = new StringBuilder();
            sb.AppendLine($"=== Tables ({tables.Count}) ===");
            sb.AppendLine();

            for (var i = 0; i < tables.Count; i++)
            {
                var table = tables[i];
                sb.AppendLine($"[{i}] Rows: {table.Rows.Count}, Columns: {table.FirstRow?.Cells?.Count ?? 0}");
                sb.AppendLine($"    Style: {table.Style?.Name ?? "(none)"}");

                // Add context: text before table (helps LLM identify the correct table)
                var precedingText = GetPrecedingText(table, 50);
                if (!string.IsNullOrWhiteSpace(precedingText))
                    sb.AppendLine($"    Preceding text: \"{precedingText}\"");

                if (includeContent)
                {
                    sb.AppendLine("    Content:");
                    for (var row = 0; row < Math.Min(3, table.Rows.Count); row++)
                    {
                        var rowText = string.Join(" | ",
                            table.Rows[row].Cells.Cast<Cell>().Select(c =>
                                c.GetText().Trim().Substring(0, Math.Min(20, c.GetText().Trim().Length))));
                        sb.AppendLine($"      {rowText}");
                    }

                    if (table.Rows.Count > 3)
                        sb.AppendLine($"      ... ({table.Rows.Count - 3} more rows)");
                }

                sb.AppendLine();
            }

            return sb.ToString();
        });
    }

    /// <summary>
    ///     Inserts a row into a table
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing tableIndex, rowIndex, optional insertBefore, rowData</param>
    /// <returns>Success message</returns>
    private Task<string> InsertRow(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var tableIndex = ArgumentHelper.GetInt(arguments, "tableIndex");
            var rowIndex = ArgumentHelper.GetInt(arguments, "rowIndex");
            var insertBefore = ArgumentHelper.GetBool(arguments, "insertBefore", false);
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);
            var dataArray = ArgumentHelper.GetArray(arguments, "rowData", false);

            var doc = new Document(path);
            if (sectionIndex >= doc.Sections.Count)
                throw new ArgumentException($"Section index {sectionIndex} out of range");

            var section = doc.Sections[sectionIndex];
            var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
            if (tableIndex < 0 || tableIndex >= tables.Count)
                throw new ArgumentException($"Table index {tableIndex} out of range");

            var table = tables[tableIndex];
            if (rowIndex < 0 || rowIndex >= table.Rows.Count)
                throw new ArgumentException($"Row index {rowIndex} out of range");

            var targetRow = table.Rows[rowIndex];
            var columnCount = targetRow.Cells.Count;
            var newRow = new Row(doc);

            for (var i = 0; i < columnCount; i++)
            {
                var newCell = new Cell(doc);
                newRow.AppendChild(newCell);
                if (dataArray != null && i < dataArray.Count)
                {
                    var cellText = dataArray[i]?.GetValue<string>() ?? "";
                    if (!string.IsNullOrEmpty(cellText))
                    {
                        var para = new Paragraph(doc);
                        // Handle multi-line text: split by \n and insert line breaks
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

            doc.Save(outputPath);
            return
                $"Successfully inserted row at index {(insertBefore ? rowIndex : rowIndex + 1)}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes a row from a table
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing tableIndex, rowIndex, optional sectionIndex</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteRow(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var tableIndex = ArgumentHelper.GetInt(arguments, "tableIndex");
            var rowIndex = ArgumentHelper.GetInt(arguments, "rowIndex");
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);

            var doc = new Document(path);
            if (sectionIndex >= doc.Sections.Count)
                throw new ArgumentException($"Section index {sectionIndex} out of range");

            var section = doc.Sections[sectionIndex];
            var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
            if (tableIndex < 0 || tableIndex >= tables.Count)
                throw new ArgumentException($"Table index {tableIndex} out of range");

            var table = tables[tableIndex];
            if (rowIndex < 0 || rowIndex >= table.Rows.Count)
                throw new ArgumentException($"Row index {rowIndex} out of range");

            var rowToDelete = table.Rows[rowIndex];
            rowToDelete.Remove();

            doc.Save(outputPath);
            return $"Successfully deleted row #{rowIndex}. Remaining rows: {table.Rows.Count}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Inserts a column into a table
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing tableIndex, columnIndex, optional insertBefore, columnData</param>
    /// <returns>Success message</returns>
    private Task<string> InsertColumn(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var tableIndex = ArgumentHelper.GetInt(arguments, "tableIndex");
            int columnIndex;
            var columnIndexNode = arguments?["columnIndex"];
            if (columnIndexNode != null)
            {
                if (columnIndexNode.GetValueKind() == JsonValueKind.String)
                {
                    var columnIndexStr = columnIndexNode.GetValue<string>();
                    if (string.IsNullOrEmpty(columnIndexStr) || !int.TryParse(columnIndexStr, out columnIndex))
                        throw new ArgumentException("columnIndex must be a valid integer");
                }
                else if (columnIndexNode.GetValueKind() == JsonValueKind.Number)
                {
                    columnIndex = columnIndexNode.GetValue<int>();
                }
                else
                {
                    throw new ArgumentException("columnIndex is required and must be an integer");
                }
            }
            else
            {
                throw new ArgumentException("columnIndex is required");
            }

            var insertBefore = ArgumentHelper.GetBool(arguments, "insertBefore", false);
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);
            var dataArray = ArgumentHelper.GetArray(arguments, "columnData", false);

            var doc = new Document(path);
            if (sectionIndex >= doc.Sections.Count)
                throw new ArgumentException($"Section index {sectionIndex} out of range");

            var section = doc.Sections[sectionIndex];
            var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
            if (tableIndex < 0 || tableIndex >= tables.Count)
                throw new ArgumentException($"Table index {tableIndex} out of range");

            var table = tables[tableIndex];
            if (table.Rows.Count == 0)
                throw new InvalidOperationException($"Table {tableIndex} has no rows");

            var firstRow = table.Rows[0];
            if (columnIndex < 0 || columnIndex >= firstRow.Cells.Count)
                throw new ArgumentException($"Column index {columnIndex} out of range");

            var insertPosition = insertBefore ? columnIndex : columnIndex + 1;

            for (var rowIdx = 0; rowIdx < table.Rows.Count; rowIdx++)
            {
                var row = table.Rows[rowIdx];
                var newCell = new Cell(doc);
                if (columnIndex < row.Cells.Count)
                {
                    var sourceCell = row.Cells[columnIndex];
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
                        // Handle multi-line text: split by \n and insert line breaks
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

            doc.Save(outputPath);
            return
                $"Successfully inserted column at index {(insertBefore ? columnIndex : columnIndex + 1)}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes a column from a table
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing tableIndex, columnIndex, optional sectionIndex</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteColumn(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var tableIndex = ArgumentHelper.GetInt(arguments, "tableIndex");
            int columnIndex;
            var columnIndexNode = arguments?["columnIndex"];
            if (columnIndexNode != null)
            {
                if (columnIndexNode.GetValueKind() == JsonValueKind.String)
                {
                    var columnIndexStr = columnIndexNode.GetValue<string>();
                    if (string.IsNullOrEmpty(columnIndexStr) || !int.TryParse(columnIndexStr, out columnIndex))
                        throw new ArgumentException("columnIndex must be a valid integer");
                }
                else if (columnIndexNode.GetValueKind() == JsonValueKind.Number)
                {
                    columnIndex = columnIndexNode.GetValue<int>();
                }
                else
                {
                    throw new ArgumentException("columnIndex is required and must be an integer");
                }
            }
            else
            {
                throw new ArgumentException("columnIndex is required");
            }

            var sectionIndex = ArgumentHelper.GetIntNullable(arguments, "sectionIndex");

            var doc = new Document(path);
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
                // If sectionIndex is not specified, search in entire document (consistent with get_tables)
                tables = doc.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
            }

            if (tableIndex < 0 || tableIndex >= tables.Count)
                throw new ArgumentException($"Table index {tableIndex} out of range");

            var table = tables[tableIndex];
            if (table.Rows.Count == 0)
                throw new InvalidOperationException($"Table {tableIndex} has no rows");

            var firstRow = table.Rows[0];
            if (columnIndex < 0 || columnIndex >= firstRow.Cells.Count)
                throw new ArgumentException($"Column index {columnIndex} out of range");

            var deletedCount = 0;
            foreach (var row in table.Rows.Cast<Row>())
                if (columnIndex < row.Cells.Count)
                {
                    row.Cells[columnIndex].Remove();
                    deletedCount++;
                }

            doc.Save(outputPath);
            return $"Successfully deleted column #{columnIndex} ({deletedCount} cells removed). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Merges cells in a table
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing tableIndex, startRow, startCol, endRow, endCol</param>
    /// <returns>Success message</returns>
    private Task<string> MergeCells(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var tableIndex = ArgumentHelper.GetInt(arguments, "tableIndex");
            var startRow = ArgumentHelper.GetInt(arguments, "startRow");
            var startCol = ArgumentHelper.GetInt(arguments, "startCol", "startColumn", "startCol or startColumn");
            var endRow = ArgumentHelper.GetInt(arguments, "endRow");
            var endCol = ArgumentHelper.GetInt(arguments, "endCol", "endColumn", "endCol or endColumn");

            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);

            var doc = new Document(path);
            if (sectionIndex >= doc.Sections.Count)
                throw new ArgumentException($"Section index {sectionIndex} out of range");

            var section = doc.Sections[sectionIndex];
            var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
            if (tableIndex < 0 || tableIndex >= tables.Count)
                throw new ArgumentException($"Table index {tableIndex} out of range");

            var table = tables[tableIndex];
            if (startRow < 0 || startRow >= table.Rows.Count || endRow < 0 || endRow >= table.Rows.Count)
                throw new ArgumentException("Row indices out of range");
            if (startRow > endRow)
                throw new ArgumentException($"Start row {startRow} cannot be greater than end row {endRow}");

            var firstRow = table.Rows[startRow];
            if (startCol < 0 || startCol >= firstRow.Cells.Count || endCol < 0 || endCol >= firstRow.Cells.Count)
                throw new ArgumentException("Column indices out of range");
            if (startCol > endCol)
                throw new ArgumentException($"Start column {startCol} cannot be greater than end column {endCol}");

            // Merge cells: Set merge flags for all cells in the range
            // The first cell (startRow, startCol) is marked as First for both horizontal and vertical merge
            // Other cells are marked as Previous to merge with the first cell
            for (var row = startRow; row <= endRow; row++)
            {
                var currentRow = table.Rows[row];
                for (var col = startCol; col <= endCol; col++)
                {
                    var cell = currentRow.Cells[col];
                    if (row == startRow && col == startCol)
                    {
                        // First cell: Set as merge origin
                        if (startRow != endRow)
                            cell.CellFormat.VerticalMerge = CellMerge.First;
                        if (startCol != endCol)
                            cell.CellFormat.HorizontalMerge = CellMerge.First;
                    }
                    else
                    {
                        // Other cells: Merge with the first cell
                        if (row == startRow)
                        {
                            // Same row as first cell: only horizontal merge
                            cell.CellFormat.HorizontalMerge = CellMerge.Previous;
                        }
                        else if (col == startCol)
                        {
                            // Same column as first cell: only vertical merge
                            cell.CellFormat.VerticalMerge = CellMerge.Previous;
                        }
                        else
                        {
                            // Both row and column differ: merge both horizontally and vertically
                            cell.CellFormat.HorizontalMerge = CellMerge.Previous;
                            cell.CellFormat.VerticalMerge = CellMerge.Previous;
                        }
                    }
                }
            }

            doc.Save(outputPath);
            return
                $"Successfully merged cells from [{startRow}, {startCol}] to [{endRow}, {endCol}]. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Splits a cell in a table
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing tableIndex, rowIndex, colIndex, optional splitRows, splitCols</param>
    /// <returns>Success message</returns>
    private Task<string> SplitCell(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var tableIndex = ArgumentHelper.GetInt(arguments, "tableIndex");

            var rowIndex = ArgumentHelper.GetInt(arguments, "rowIndex");
            var colIndex = ArgumentHelper.GetInt(arguments, "colIndex", "columnIndex", "colIndex or columnIndex");
            var splitRows = ArgumentHelper.GetInt(arguments, "splitRows", 2);
            var splitCols = ArgumentHelper.GetInt(arguments, "splitCols", 2);
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);

            var doc = new Document(path);
            if (sectionIndex >= doc.Sections.Count)
                throw new ArgumentException($"Section index {sectionIndex} out of range");

            var section = doc.Sections[sectionIndex];
            var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
            if (tableIndex < 0 || tableIndex >= tables.Count)
                throw new ArgumentException($"Table index {tableIndex} out of range");

            var table = tables[tableIndex];
            if (rowIndex < 0 || rowIndex >= table.Rows.Count)
                throw new ArgumentException($"Row index {rowIndex} out of range");

            var row = table.Rows[rowIndex];
            if (colIndex < 0 || colIndex >= row.Cells.Count)
                throw new ArgumentException($"Column index {colIndex} out of range");

            var cell = row.Cells[colIndex];
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
                    var insertAfterRowIndex = rowIndex + r - 1;
                    if (insertAfterRowIndex < table.Rows.Count)
                    {
                        var newRow = new Row(doc);
                        var sourceRow = table.Rows[rowIndex];
                        for (var c = 0; c < sourceRow.Cells.Count; c++)
                        {
                            var newCell = new Cell(doc);
                            var sourceCell = sourceRow.Cells[c];
                            newCell.CellFormat.Width = sourceCell.CellFormat.Width;
                            newCell.CellFormat.VerticalAlignment = sourceCell.CellFormat.VerticalAlignment;
                            newCell.CellFormat.SetPaddings(
                                sourceCell.CellFormat.TopPadding,
                                sourceCell.CellFormat.BottomPadding,
                                sourceCell.CellFormat.LeftPadding,
                                sourceCell.CellFormat.RightPadding
                            );
                            var para = new Paragraph(doc);
                            newCell.AppendChild(para);
                            newRow.AppendChild(newCell);
                        }

                        table.InsertAfter(newRow, table.Rows[insertAfterRowIndex]);
                    }
                }

            doc.Save(outputPath);
            return
                $"Successfully split cell [{rowIndex}, {colIndex}] into {splitRows} rows x {splitCols} columns. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits cell format properties
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing tableIndex, optional rowIndex, colIndex, formatting options</param>
    /// <returns>Success message</returns>
    private Task<string> EditCellFormat(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var tableIndex = ArgumentHelper.GetInt(arguments, "tableIndex");
            var applyToRow = ArgumentHelper.GetBool(arguments, "applyToRow", false);
            var applyToColumn = ArgumentHelper.GetBool(arguments, "applyToColumn", false);
            var applyToTable = ArgumentHelper.GetBool(arguments, "applyToTable", false);

            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);

            var doc = new Document(path);
            if (sectionIndex >= doc.Sections.Count)
                throw new ArgumentException($"Section index {sectionIndex} out of range");

            var section = doc.Sections[sectionIndex];
            var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
            if (tableIndex < 0 || tableIndex >= tables.Count)
                throw new ArgumentException($"Table index {tableIndex} out of range");

            var table = tables[tableIndex];

            // Determine target cells based on apply flags
            var targetCells = new List<Cell>();

            if (applyToTable)
            {
                foreach (var row in table.Rows.Cast<Row>())
                    targetCells.AddRange(row.Cells.Cast<Cell>());
            }
            else if (applyToRow)
            {
                var rowIndex = ArgumentHelper.GetInt(arguments, "rowIndex");
                if (rowIndex < 0 || rowIndex >= table.Rows.Count)
                    throw new ArgumentException($"Row index {rowIndex} out of range");
                targetCells.AddRange(table.Rows[rowIndex].Cells.Cast<Cell>());
            }
            else if (applyToColumn)
            {
                var colIndex = ArgumentHelper.GetInt(arguments, "colIndex", "columnIndex", "colIndex or columnIndex");
                foreach (var row in table.Rows.Cast<Row>())
                    if (colIndex < row.Cells.Count)
                        targetCells.Add(row.Cells[colIndex]);
            }
            else
            {
                // Single cell (original behavior)
                var rowIndex = ArgumentHelper.GetInt(arguments, "rowIndex");
                var colIndex = ArgumentHelper.GetInt(arguments, "colIndex", "columnIndex", "colIndex or columnIndex");
                if (rowIndex < 0 || rowIndex >= table.Rows.Count)
                    throw new ArgumentException($"Row index {rowIndex} out of range");
                var row = table.Rows[rowIndex];
                if (colIndex < 0 || colIndex >= row.Cells.Count)
                    throw new ArgumentException($"Column index {colIndex} out of range");
                targetCells.Add(row.Cells[colIndex]);
            }

            if (targetCells.Count == 0)
                throw new ArgumentException("No target cells found");

            var backgroundColor = ArgumentHelper.GetStringNullable(arguments, "backgroundColor");
            var alignment = ArgumentHelper.GetStringNullable(arguments, "alignment");
            var verticalAlignment = ArgumentHelper.GetStringNullable(arguments, "verticalAlignment");
            var paddingTop = ArgumentHelper.GetDoubleNullable(arguments, "paddingTop");
            var paddingBottom = ArgumentHelper.GetDoubleNullable(arguments, "paddingBottom");
            var paddingLeft = ArgumentHelper.GetDoubleNullable(arguments, "paddingLeft");
            var paddingRight = ArgumentHelper.GetDoubleNullable(arguments, "paddingRight");
            var fontName = ArgumentHelper.GetStringNullable(arguments, "fontName");
            var fontNameAscii = ArgumentHelper.GetStringNullable(arguments, "fontNameAscii");
            var fontNameFarEast = ArgumentHelper.GetStringNullable(arguments, "fontNameFarEast");
            var fontSize = ArgumentHelper.GetDoubleNullable(arguments, "fontSize");
            var bold = ArgumentHelper.GetBoolNullable(arguments, "bold");
            var italic = ArgumentHelper.GetBoolNullable(arguments, "italic");
            var color = ArgumentHelper.GetStringNullable(arguments, "color");

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

                if (!string.IsNullOrEmpty(verticalAlignment))
                    cellFormat.VerticalAlignment = verticalAlignment.ToLower() switch
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
                    ? $"row {ArgumentHelper.GetInt(arguments, "rowIndex")}"
                    : applyToColumn
                        ? $"column {ArgumentHelper.GetInt(arguments, "colIndex", "columnIndex", "colIndex or columnIndex")}"
                        : $"cell [{ArgumentHelper.GetInt(arguments, "rowIndex")}, {ArgumentHelper.GetInt(arguments, "colIndex", "columnIndex", "colIndex or columnIndex")}]";

            doc.Save(outputPath);
            return $"Successfully edited {targetDescription} format ({targetCells.Count} cells). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Moves a table to a different position
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing tableIndex, targetParagraphIndex, optional sectionIndex</param>
    /// <returns>Success message</returns>
    private Task<string> MoveTable(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var tableIndex = ArgumentHelper.GetInt(arguments, "tableIndex");
            var targetParagraphIndex = ArgumentHelper.GetInt(arguments, "targetParagraphIndex");
            var sectionIndex = ArgumentHelper.GetIntNullable(arguments, "sectionIndex");

            var doc = new Document(path);
            var sectionIdx = sectionIndex ?? 0;
            if (sectionIdx < 0 || sectionIdx >= doc.Sections.Count)
                throw new ArgumentException($"sectionIndex must be between 0 and {doc.Sections.Count - 1}");

            var section = doc.Sections[sectionIdx];
            var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
            // Use false to get only direct child paragraphs of Body (exclude paragraphs inside tables/shapes)
            var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, false).Cast<Paragraph>().ToList();

            if (tableIndex < 0 || tableIndex >= tables.Count)
                throw new ArgumentException($"tableIndex must be between 0 and {tables.Count - 1}");

            var table = tables[tableIndex];
            Paragraph? targetPara;

            if (targetParagraphIndex == -1)
            {
                // targetParagraphIndex=-1 means document end - use last paragraph in Body
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

            doc.Save(outputPath);
            return $"Successfully moved table {tableIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Copies a table to another location
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing sourceTableIndex, targetParagraphIndex, optional section indices</param>
    /// <returns>Success message</returns>
    private Task<string> CopyTable(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            // Accept both sourceTableIndex and tableIndex for compatibility
            var sourceTableIndex =
                ArgumentHelper.GetInt(arguments, "sourceTableIndex", "tableIndex", "sourceTableIndex or tableIndex");
            var targetParagraphIndex = ArgumentHelper.GetInt(arguments, "targetParagraphIndex");
            var sourceSectionIndex = ArgumentHelper.GetIntNullable(arguments, "sourceSectionIndex");
            var targetSectionIndex = ArgumentHelper.GetIntNullable(arguments, "targetSectionIndex");

            var doc = new Document(path);
            var sourceSectionIdx = sourceSectionIndex ?? 0;
            var targetSectionIdx = targetSectionIndex ?? 0;

            if (sourceSectionIdx < 0 || sourceSectionIdx >= doc.Sections.Count)
                throw new ArgumentException($"sourceSectionIndex must be between 0 and {doc.Sections.Count - 1}");
            if (targetSectionIdx < 0 || targetSectionIdx >= doc.Sections.Count)
                throw new ArgumentException($"targetSectionIndex must be between 0 and {doc.Sections.Count - 1}");

            var sourceSection = doc.Sections[sourceSectionIdx];
            var sourceTables = sourceSection.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
            if (sourceTableIndex < 0 || sourceTableIndex >= sourceTables.Count)
                throw new ArgumentException($"sourceTableIndex must be between 0 and {sourceTables.Count - 1}");

            var sourceTable = sourceTables[sourceTableIndex];
            var targetSection = doc.Sections[targetSectionIdx];
            var targetParagraphs =
                targetSection.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();

            Paragraph? targetPara;
            if (targetParagraphIndex == -1)
            {
                // targetParagraphIndex=-1 means document end - use last paragraph
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

            // InsertAfter requires the reference node to be a direct child of the parent
            Node? insertionPoint;

            if (targetPara.ParentNode == targetSection.Body)
            {
                insertionPoint = targetPara;
            }
            else
            {
                // Find the direct child paragraph in Body that contains or is after targetPara
                var bodyParagraphs = targetSection.Body.GetChildNodes(NodeType.Paragraph, false);
                Paragraph? directPara = null;

                foreach (var para in bodyParagraphs.Cast<Paragraph>())
                    if (para == targetPara)
                    {
                        directPara = para;
                        break;
                    }

                // If not found, use the last direct child paragraph as insertion point
                if (directPara == null && bodyParagraphs.Count > 0)
                    directPara = bodyParagraphs[^1] as Paragraph;

                insertionPoint = directPara ?? targetSection.Body.LastChild;
            }

            if (insertionPoint == null)
                throw new ArgumentException(
                    $"Unable to find valid insertion point (targetParagraphIndex: {targetParagraphIndex})");

            var clonedTable = (Table)sourceTable.Clone(true);
            targetSection.Body.InsertAfter(clonedTable, insertionPoint);

            doc.Save(outputPath);
            return
                $"Successfully copied table {sourceTableIndex} to paragraph {targetParagraphIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets detailed table structure information
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="arguments">JSON arguments containing tableIndex, optional includeContent, includeCellFormatting</param>
    /// <returns>Formatted string with table structure</returns>
    private Task<string> GetTableStructure(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var tableIndex = ArgumentHelper.GetInt(arguments, "tableIndex");
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);
            var includeContent = ArgumentHelper.GetBool(arguments, "includeContent", false);
            var includeCellFormatting = ArgumentHelper.GetBool(arguments, "includeCellFormatting", true);

            var doc = new Document(path);
            if (sectionIndex >= doc.Sections.Count)
                throw new ArgumentException($"Section index {sectionIndex} out of range");

            var section = doc.Sections[sectionIndex];
            var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
            if (tableIndex < 0 || tableIndex >= tables.Count)
                throw new ArgumentException($"Table index {tableIndex} out of range");

            var table = tables[tableIndex];
            var result = new StringBuilder();

            result.AppendLine($"=== Table #{tableIndex} Structure ===\n");
            result.AppendLine("【Basic Info】");
            result.AppendLine($"Rows: {table.Rows.Count}");
            if (table.Rows.Count > 0)
                result.AppendLine($"Columns: {table.Rows[0].Cells.Count}");
            result.AppendLine();

            result.AppendLine("【Table Format】");
            result.AppendLine($"Alignment: {table.Alignment}");
            result.AppendLine($"Style: {table.Style?.Name ?? "None"}");
            result.AppendLine($"Left Indent: {table.LeftIndent:F2} pt");
            if (table.PreferredWidth.Type != PreferredWidthType.Auto)
                result.AppendLine($"Width: {table.PreferredWidth.Value} ({table.PreferredWidth.Type})");
            result.AppendLine($"Allow Auto Fit: {table.AllowAutoFit}");
            result.AppendLine();

            if (includeContent)
            {
                result.AppendLine("【Content Preview】");
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
                result.AppendLine("【First Cell Formatting】");
                var cell = table.Rows[0].Cells[0];
                result.AppendLine($"Top Padding: {cell.CellFormat.TopPadding:F2} pt");
                result.AppendLine($"Bottom Padding: {cell.CellFormat.BottomPadding:F2} pt");
                result.AppendLine($"Left Padding: {cell.CellFormat.LeftPadding:F2} pt");
                result.AppendLine($"Right Padding: {cell.CellFormat.RightPadding:F2} pt");
                result.AppendLine($"Vertical Alignment: {cell.CellFormat.VerticalAlignment}");
                result.AppendLine();
            }

            return result.ToString();
        });
    }

    /// <summary>
    ///     Sets table border properties
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing tableIndex, optional border flags, lineStyle, lineWidth, lineColor</param>
    /// <returns>Success message</returns>
    private Task<string> SetTableBorder(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var tableIndex = ArgumentHelper.GetInt(arguments, "tableIndex");
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);
            var rowIndex = ArgumentHelper.GetIntNullable(arguments, "rowIndex");
            var columnIndex = ArgumentHelper.GetIntNullable(arguments, "columnIndex");

            var doc = new Document(path);
            if (sectionIndex >= doc.Sections.Count)
                throw new ArgumentException($"Section index {sectionIndex} out of range");

            var section = doc.Sections[sectionIndex];
            var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
            if (tableIndex >= tables.Count)
                throw new ArgumentException($"Table index {tableIndex} out of range");

            var table = tables[tableIndex];
            var defaultLineStyle = ArgumentHelper.GetString(arguments, "lineStyle", "single");
            var defaultLineWidth = ArgumentHelper.GetDouble(arguments, "lineWidth", "lineWidth", false, 0.5);
            var defaultLineColor = ArgumentHelper.GetString(arguments, "lineColor", "000000");

            var lineStyle = GetLineStyle(defaultLineStyle);
            var lineWidth = defaultLineWidth;
            var lineColor = ColorHelper.ParseColor(defaultLineColor);

            var targetCells = new List<Cell>();
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
                if (ArgumentHelper.GetBool(arguments, "borderTop", false))
                {
                    borders.Top.LineStyle = lineStyle;
                    borders.Top.LineWidth = lineWidth;
                    borders.Top.Color = lineColor;
                }
                else if (arguments?["borderTop"] != null)
                {
                    borders.Top.LineStyle = LineStyle.None;
                }

                if (ArgumentHelper.GetBool(arguments, "borderBottom", false))
                {
                    borders.Bottom.LineStyle = lineStyle;
                    borders.Bottom.LineWidth = lineWidth;
                    borders.Bottom.Color = lineColor;
                }
                else if (arguments?["borderBottom"] != null)
                {
                    borders.Bottom.LineStyle = LineStyle.None;
                }

                if (ArgumentHelper.GetBool(arguments, "borderLeft", false))
                {
                    borders.Left.LineStyle = lineStyle;
                    borders.Left.LineWidth = lineWidth;
                    borders.Left.Color = lineColor;
                }
                else if (arguments?["borderLeft"] != null)
                {
                    borders.Left.LineStyle = LineStyle.None;
                }

                if (ArgumentHelper.GetBool(arguments, "borderRight", false))
                {
                    borders.Right.LineStyle = lineStyle;
                    borders.Right.LineWidth = lineWidth;
                    borders.Right.Color = lineColor;
                }
                else if (arguments?["borderRight"] != null)
                {
                    borders.Right.LineStyle = LineStyle.None;
                }
            }

            doc.Save(outputPath);
            return $"Successfully set table {tableIndex} borders. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets column width for a table
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing tableIndex, columnIndex, columnWidth</param>
    /// <returns>Success message</returns>
    private Task<string> SetColumnWidth(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var tableIndex = ArgumentHelper.GetInt(arguments, "tableIndex");
            var columnIndex = ArgumentHelper.GetInt(arguments, "columnIndex");
            var columnWidth = ArgumentHelper.GetDouble(arguments, "columnWidth");

            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);

            if (columnWidth <= 0)
                throw new ArgumentException($"Column width {columnWidth} must be greater than 0");

            var doc = new Document(path);
            if (sectionIndex >= doc.Sections.Count)
                throw new ArgumentException($"Section index {sectionIndex} out of range");

            var section = doc.Sections[sectionIndex];
            var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
            if (tableIndex < 0 || tableIndex >= tables.Count)
                throw new ArgumentException($"Table index {tableIndex} out of range");

            var table = tables[tableIndex];
            if (table.Rows.Count == 0)
                throw new InvalidOperationException($"Table {tableIndex} has no rows");

            var firstRow = table.Rows[0];
            if (columnIndex < 0 || columnIndex >= firstRow.Cells.Count)
                throw new ArgumentException($"Column index {columnIndex} out of range");

            var cellsUpdated = 0;
            foreach (var row in table.Rows.Cast<Row>())
                if (columnIndex < row.Cells.Count)
                {
                    row.Cells[columnIndex].CellFormat.PreferredWidth = PreferredWidth.FromPoints(columnWidth);
                    cellsUpdated++;
                }

            doc.Save(outputPath);
            return
                $"Successfully set column {columnIndex} width to {columnWidth} pt ({cellsUpdated} cells updated). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets row height for a table
    /// </summary>
    /// <param name="path">Document file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing tableIndex, rowIndex, rowHeight, optional heightRule</param>
    /// <returns>Success message</returns>
    private Task<string> SetRowHeight(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var tableIndex = ArgumentHelper.GetInt(arguments, "tableIndex");
            var rowIndex = ArgumentHelper.GetInt(arguments, "rowIndex");
            var rowHeight = ArgumentHelper.GetDouble(arguments, "rowHeight");

            var heightRule = ArgumentHelper.GetString(arguments, "heightRule", "atLeast");
            var sectionIndex = ArgumentHelper.GetInt(arguments, "sectionIndex", 0);

            if (rowHeight <= 0)
                throw new ArgumentException($"Row height {rowHeight} must be greater than 0");

            var doc = new Document(path);
            if (sectionIndex >= doc.Sections.Count)
                throw new ArgumentException($"Section index {sectionIndex} out of range");

            var section = doc.Sections[sectionIndex];
            var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Table>().ToList();
            if (tableIndex < 0 || tableIndex >= tables.Count)
                throw new ArgumentException($"Table index {tableIndex} out of range");

            var table = tables[tableIndex];
            if (rowIndex < 0 || rowIndex >= table.Rows.Count)
                throw new ArgumentException($"Row index {rowIndex} out of range");

            var row = table.Rows[rowIndex];
            row.RowFormat.HeightRule = heightRule.ToLower() switch
            {
                "auto" => HeightRule.Auto,
                "atLeast" => HeightRule.AtLeast,
                "exactly" => HeightRule.Exactly,
                _ => HeightRule.AtLeast
            };
            row.RowFormat.Height = rowHeight;

            doc.Save(outputPath);
            return $"Successfully set row {rowIndex} height to {rowHeight} pt ({heightRule}). Output: {outputPath}";
        });
    }

    private Dictionary<int, string> ParseColorDictionary(JsonNode? node)
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

    private List<(int row, int col, string color)> ParseCellColors(JsonNode? node)
    {
        var result = new List<(int row, int col, string color)>();
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

    private List<(int startRow, int endRow, int startCol, int endCol)> ParseMergeCells(JsonNode? node)
    {
        var result = new List<(int startRow, int endRow, int startCol, int endCol)>();
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

    private CellVerticalAlignment GetVerticalAlignment(string alignment)
    {
        return alignment.ToLower() switch
        {
            "top" => CellVerticalAlignment.Top,
            "bottom" => CellVerticalAlignment.Bottom,
            _ => CellVerticalAlignment.Center
        };
    }

    private LineStyle GetLineStyle(string style)
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
    ///     Gets the text content preceding a table node for context identification
    /// </summary>
    /// <param name="table">The table node</param>
    /// <param name="maxLength">Maximum characters to return</param>
    /// <returns>Preceding text trimmed to maxLength, or empty string if none found</returns>
    private static string GetPrecedingText(Table table, int maxLength)
    {
        var precedingSibling = table.PreviousSibling;
        while (precedingSibling != null)
        {
            if (precedingSibling is Paragraph para)
            {
                var text = para.GetText().Trim();
                // Skip empty paragraphs and section breaks
                if (!string.IsNullOrWhiteSpace(text) && !text.StartsWith("\f"))
                {
                    // Clean control characters
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