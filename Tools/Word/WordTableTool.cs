using System.ComponentModel;
using Aspose.Words;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using AsposeMcpServer.Handlers.Word.Table;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Word;

/// <summary>
///     Unified tool for managing tables in Word documents.
///     Dispatches operations to specialized handlers.
/// </summary>
[McpServerToolType]
public class WordTableTool
{
    /// <summary>
    ///     Handler registry for Word table operations.
    /// </summary>
    private readonly HandlerRegistry<Document> _handlerRegistry;

    /// <summary>
    ///     Identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document session operations.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the WordTableTool class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document operations.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public WordTableTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
        _handlerRegistry = WordTableHandlerRegistry.Create();
    }

    /// <summary>
    ///     Executes a Word table operation (create, delete, get, insert_row, delete_row, insert_column, delete_column,
    ///     merge_cells, split_cell, edit_cell_format, move_table, copy_table, get_structure, set_border, set_column_width,
    ///     set_row_height).
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: create, delete, get, insert_row, delete_row, insert_column,
    ///     delete_column, merge_cells, split_cell, edit_cell_format, move_table, copy_table, get_structure, set_border,
    ///     set_column_width, set_row_height.
    /// </param>
    /// <param name="path">Word document file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (file mode only).</param>
    /// <param name="tableIndex">Table index (0-based).</param>
    /// <param name="sectionIndex">Section index (0-based).</param>
    /// <param name="rows">Number of rows (for create).</param>
    /// <param name="columns">Number of columns (for create).</param>
    /// <param name="paragraphIndex">Paragraph index to insert after (-1 for end, for create).</param>
    /// <param name="tableData">Table data as JSON 2D array (for create).</param>
    /// <param name="tableWidth">Table width in points (for create).</param>
    /// <param name="autoFit">Auto-fit table (for create, default: true).</param>
    /// <param name="hasHeader">Header row with alternating colors (for create, default: true).</param>
    /// <param name="headerBackgroundColor">Header background color hex (for create).</param>
    /// <param name="cellBackgroundColor">Cell background color hex (for create).</param>
    /// <param name="alternatingRowColor">Alternating row color hex (for create).</param>
    /// <param name="rowColors">Row colors by index as JSON object (for create).</param>
    /// <param name="cellColors">Cell colors as JSON array (for create).</param>
    /// <param name="mergeCells">Cells to merge as JSON array (for create).</param>
    /// <param name="fontName">Font name (for create).</param>
    /// <param name="fontSize">Font size in points (for create).</param>
    /// <param name="rowIndex">Row index (0-based, for insert_row/delete_row/set_row_height).</param>
    /// <param name="columnIndex">Column index (0-based, for insert_column/delete_column/set_column_width).</param>
    /// <param name="verticalAlignment">Vertical alignment (for create).</param>
    /// <param name="rowData">Row data as JSON array (for insert_row).</param>
    /// <param name="columnData">Column data as JSON array (for insert_column).</param>
    /// <param name="insertBefore">Insert before target position (for insert_row/insert_column).</param>
    /// <param name="startRow">Start row for merge (for merge_cells).</param>
    /// <param name="endRow">End row for merge (for merge_cells).</param>
    /// <param name="startCol">Start column for merge (for merge_cells).</param>
    /// <param name="endCol">End column for merge (for merge_cells).</param>
    /// <param name="splitRows">Number of rows to split into (for split_cell).</param>
    /// <param name="splitCols">Number of columns to split into (for split_cell).</param>
    /// <param name="applyToRow">Apply formatting to entire row (for edit_cell_format).</param>
    /// <param name="applyToColumn">Apply formatting to entire column (for edit_cell_format).</param>
    /// <param name="applyToTable">Apply formatting to entire table (for edit_cell_format).</param>
    /// <param name="backgroundColor">Background color hex (for edit_cell_format).</param>
    /// <param name="alignment">Text alignment (for edit_cell_format).</param>
    /// <param name="verticalAlignmentFormat">Vertical alignment for cells (for edit_cell_format).</param>
    /// <param name="paddingTop">Top padding in points (for edit_cell_format).</param>
    /// <param name="paddingBottom">Bottom padding in points (for edit_cell_format).</param>
    /// <param name="paddingLeft">Left padding in points (for edit_cell_format).</param>
    /// <param name="paddingRight">Right padding in points (for edit_cell_format).</param>
    /// <param name="fontNameAscii">Font name for ASCII (for edit_cell_format).</param>
    /// <param name="fontNameFarEast">Font name for Far East (for edit_cell_format).</param>
    /// <param name="cellFontSize">Font size for cells in points (for edit_cell_format).</param>
    /// <param name="bold">Bold text (for edit_cell_format).</param>
    /// <param name="italic">Italic text (for edit_cell_format).</param>
    /// <param name="color">Text color hex (for edit_cell_format).</param>
    /// <param name="targetParagraphIndex">Target paragraph index for move/copy.</param>
    /// <param name="sourceSectionIndex">Source section index for copy.</param>
    /// <param name="targetSectionIndex">Target section index for copy.</param>
    /// <param name="includeContent">Include content in get_structure.</param>
    /// <param name="includeCellFormatting">Include cell formatting in get_structure.</param>
    /// <param name="borderTop">Enable top border (for set_border).</param>
    /// <param name="borderBottom">Enable bottom border (for set_border).</param>
    /// <param name="borderLeft">Enable left border (for set_border).</param>
    /// <param name="borderRight">Enable right border (for set_border).</param>
    /// <param name="lineStyle">Border line style (for set_border).</param>
    /// <param name="lineWidth">Border line width in points (for set_border).</param>
    /// <param name="lineColor">Border line color hex (for set_border).</param>
    /// <param name="columnWidth">Column width in points (for set_column_width).</param>
    /// <param name="rowHeight">Row height in points (for set_row_height).</param>
    /// <param name="heightRule">Height rule (for set_row_height).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
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
        var effectiveOutputPath = outputPath ?? path;
        if (!string.IsNullOrEmpty(effectiveOutputPath))
            SecurityHelper.ValidateFilePath(effectiveOutputPath, "outputPath", true);

        var parameters = BuildParameters(tableIndex, sectionIndex, rows, columns, paragraphIndex, tableData,
            tableWidth, autoFit, hasHeader, headerBackgroundColor, cellBackgroundColor, alternatingRowColor,
            rowColors, cellColors, mergeCells, fontName, fontSize, verticalAlignment, rowIndex, columnIndex,
            rowData, columnData, insertBefore, startRow, endRow, startCol, endCol, splitRows, splitCols,
            applyToRow, applyToColumn, applyToTable, backgroundColor, alignment, verticalAlignmentFormat,
            paddingTop, paddingBottom, paddingLeft, paddingRight, fontNameAscii, fontNameFarEast, cellFontSize,
            bold, italic, color, targetParagraphIndex, sourceSectionIndex, targetSectionIndex, includeContent,
            includeCellFormatting, borderTop, borderBottom, borderLeft, borderRight, lineStyle, lineWidth,
            lineColor, columnWidth, rowHeight, heightRule);

        var handler = _handlerRegistry.GetHandler(operation);

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path, _identityAccessor);

        var operationContext = new OperationContext<Document>
        {
            Document = ctx.Document,
            SessionManager = _sessionManager,
            IdentityAccessor = _identityAccessor,
            SessionId = sessionId,
            SourcePath = path,
            OutputPath = effectiveOutputPath
        };

        var result = handler.Execute(operationContext, parameters);

        if (operationContext.IsModified)
            ctx.Save(effectiveOutputPath);

        return AppendOutputMessage(result, ctx, effectiveOutputPath);
    }

    /// <summary>
    ///     Builds the operation parameters from input values.
    /// </summary>
    private static OperationParameters BuildParameters(
        int tableIndex, int? sectionIndex, int? rows, int? columns, int paragraphIndex, string? tableData,
        double? tableWidth, bool autoFit, bool hasHeader, string? headerBackgroundColor, string? cellBackgroundColor,
        string? alternatingRowColor, string? rowColors, string? cellColors, string? mergeCells, string? fontName,
        double? fontSize, string verticalAlignment, int? rowIndex, int? columnIndex, string? rowData,
        string? columnData, bool insertBefore, int? startRow, int? endRow, int? startCol, int? endCol,
        int splitRows, int splitCols, bool applyToRow, bool applyToColumn, bool applyToTable,
        string? backgroundColor, string? alignment, string? verticalAlignmentFormat, double? paddingTop,
        double? paddingBottom, double? paddingLeft, double? paddingRight, string? fontNameAscii,
        string? fontNameFarEast, double? cellFontSize, bool? bold, bool? italic, string? color,
        int targetParagraphIndex, int? sourceSectionIndex, int? targetSectionIndex, bool includeContent,
        bool includeCellFormatting, bool borderTop, bool borderBottom, bool borderLeft, bool borderRight,
        string lineStyle, double lineWidth, string lineColor, double? columnWidth, double? rowHeight,
        string heightRule)
    {
        var parameters = new OperationParameters();

        parameters.Set("tableIndex", tableIndex);
        parameters.Set("sectionIndex", sectionIndex);
        parameters.Set("rows", rows);
        parameters.Set("columns", columns);
        parameters.Set("paragraphIndex", paragraphIndex);
        parameters.Set("tableData", tableData);
        parameters.Set("tableWidth", tableWidth);
        parameters.Set("autoFit", autoFit);
        parameters.Set("hasHeader", hasHeader);
        parameters.Set("headerBackgroundColor", headerBackgroundColor);
        parameters.Set("cellBackgroundColor", cellBackgroundColor);
        parameters.Set("alternatingRowColor", alternatingRowColor);
        parameters.Set("rowColors", rowColors);
        parameters.Set("cellColors", cellColors);
        parameters.Set("mergeCells", mergeCells);
        parameters.Set("fontName", fontName);
        parameters.Set("fontSize", fontSize);
        parameters.Set("verticalAlignment", verticalAlignment);
        parameters.Set("rowIndex", rowIndex);
        parameters.Set("columnIndex", columnIndex);
        parameters.Set("rowData", rowData);
        parameters.Set("columnData", columnData);
        parameters.Set("insertBefore", insertBefore);
        parameters.Set("startRow", startRow);
        parameters.Set("endRow", endRow);
        parameters.Set("startCol", startCol);
        parameters.Set("endCol", endCol);
        parameters.Set("splitRows", splitRows);
        parameters.Set("splitCols", splitCols);
        parameters.Set("applyToRow", applyToRow);
        parameters.Set("applyToColumn", applyToColumn);
        parameters.Set("applyToTable", applyToTable);
        parameters.Set("backgroundColor", backgroundColor);
        parameters.Set("alignment", alignment);
        parameters.Set("verticalAlignmentFormat", verticalAlignmentFormat);
        parameters.Set("paddingTop", paddingTop);
        parameters.Set("paddingBottom", paddingBottom);
        parameters.Set("paddingLeft", paddingLeft);
        parameters.Set("paddingRight", paddingRight);
        parameters.Set("fontNameAscii", fontNameAscii);
        parameters.Set("fontNameFarEast", fontNameFarEast);
        parameters.Set("cellFontSize", cellFontSize);
        parameters.Set("bold", bold);
        parameters.Set("italic", italic);
        parameters.Set("color", color);
        parameters.Set("targetParagraphIndex", targetParagraphIndex);
        parameters.Set("sourceSectionIndex", sourceSectionIndex);
        parameters.Set("targetSectionIndex", targetSectionIndex);
        parameters.Set("includeContent", includeContent);
        parameters.Set("includeCellFormatting", includeCellFormatting);
        parameters.Set("borderTop", borderTop);
        parameters.Set("borderBottom", borderBottom);
        parameters.Set("borderLeft", borderLeft);
        parameters.Set("borderRight", borderRight);
        parameters.Set("lineStyle", lineStyle);
        parameters.Set("lineWidth", lineWidth);
        parameters.Set("lineColor", lineColor);
        parameters.Set("columnWidth", columnWidth);
        parameters.Set("rowHeight", rowHeight);
        parameters.Set("heightRule", heightRule);

        return parameters;
    }

    /// <summary>
    ///     Appends output message to result.
    /// </summary>
    private static string AppendOutputMessage(string result, DocumentContext<Document> ctx, string? outputPath)
    {
        var outputMessage = ctx.GetOutputMessage(outputPath);
        if (!string.IsNullOrEmpty(outputMessage) && !result.Contains(outputMessage))
            return result + "\n" + outputMessage;
        return result;
    }
}
