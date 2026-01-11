using Aspose.Words;
using Aspose.Words.Tables;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Table;

/// <summary>
///     Handler for editing cell format in Word document tables.
/// </summary>
public class EditCellFormatWordTableHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "edit_cell_format";

    /// <summary>
    ///     Edits cell format properties.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: tableIndex, rowIndex, columnIndex, applyToRow, applyToColumn, applyToTable,
    ///     backgroundColor, alignment, verticalAlignmentFormat, paddingTop, paddingBottom, paddingLeft, paddingRight,
    ///     fontName, fontNameAscii, fontNameFarEast, cellFontSize, bold, italic, color, sectionIndex.
    /// </param>
    /// <returns>Success message with formatted cell count.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or indices are out of range.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var tableIndex = parameters.GetOptional("tableIndex", 0);
        var rowIndex = parameters.GetOptional<int?>("rowIndex");
        var columnIndex = parameters.GetOptional<int?>("columnIndex");
        var applyToRow = parameters.GetOptional("applyToRow", false);
        var applyToColumn = parameters.GetOptional("applyToColumn", false);
        var applyToTable = parameters.GetOptional("applyToTable", false);
        var backgroundColor = parameters.GetOptional<string?>("backgroundColor");
        var alignment = parameters.GetOptional<string?>("alignment");
        var verticalAlignmentStr = parameters.GetOptional<string?>("verticalAlignmentFormat");
        var paddingTop = parameters.GetOptional<double?>("paddingTop");
        var paddingBottom = parameters.GetOptional<double?>("paddingBottom");
        var paddingLeft = parameters.GetOptional<double?>("paddingLeft");
        var paddingRight = parameters.GetOptional<double?>("paddingRight");
        var fontName = parameters.GetOptional<string?>("fontName");
        var fontNameAscii = parameters.GetOptional<string?>("fontNameAscii");
        var fontNameFarEast = parameters.GetOptional<string?>("fontNameFarEast");
        var fontSize = parameters.GetOptional<double?>("cellFontSize");
        var bold = parameters.GetOptional<bool?>("bold");
        var italic = parameters.GetOptional<bool?>("italic");
        var color = parameters.GetOptional<string?>("color");
        var sectionIndex = parameters.GetOptional<int?>("sectionIndex");

        var doc = context.Document;
        var actualSectionIndex = sectionIndex ?? 0;
        if (actualSectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"Section index {actualSectionIndex} out of range");

        var section = doc.Sections[actualSectionIndex];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Aspose.Words.Tables.Table>().ToList();
        if (tableIndex < 0 || tableIndex >= tables.Count)
            throw new ArgumentException($"Table index {tableIndex} out of range");

        var table = tables[tableIndex];

        var targetCells = GetTargetCells(table, rowIndex, columnIndex, applyToRow, applyToColumn, applyToTable);

        if (targetCells.Count == 0)
            throw new ArgumentException("No target cells found");

        var hasTextFormatting = !string.IsNullOrEmpty(fontName) || !string.IsNullOrEmpty(fontNameAscii) ||
                                !string.IsNullOrEmpty(fontNameFarEast) || fontSize.HasValue ||
                                bold.HasValue || italic.HasValue || !string.IsNullOrEmpty(color);

        foreach (var cell in targetCells)
        {
            ApplyCellFormatting(cell, backgroundColor, alignment, verticalAlignmentStr,
                paddingTop, paddingBottom, paddingLeft, paddingRight);

            if (hasTextFormatting)
                ApplyTextFormatting(cell, fontName, fontNameAscii, fontNameFarEast, fontSize, bold, italic, color);
        }

        var targetDescription = GetTargetDescription(applyToTable, applyToRow, applyToColumn, rowIndex, columnIndex);

        MarkModified(context);

        return Success($"Successfully edited {targetDescription} format ({targetCells.Count} cells).");
    }

    /// <summary>
    ///     Gets the target cells based on the specified parameters.
    /// </summary>
    /// <param name="table">The table.</param>
    /// <param name="rowIndex">The row index.</param>
    /// <param name="columnIndex">The column index.</param>
    /// <param name="applyToRow">Whether to apply to entire row.</param>
    /// <param name="applyToColumn">Whether to apply to entire column.</param>
    /// <param name="applyToTable">Whether to apply to entire table.</param>
    /// <returns>List of target cells.</returns>
    /// <exception cref="ArgumentException">Thrown when required indices are missing.</exception>
    private static List<Cell> GetTargetCells(Aspose.Words.Tables.Table table, int? rowIndex, int? columnIndex,
        bool applyToRow, bool applyToColumn, bool applyToTable)
    {
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

        return targetCells;
    }

    /// <summary>
    ///     Applies cell formatting to a cell.
    /// </summary>
    /// <param name="cell">The cell.</param>
    /// <param name="backgroundColor">Background color.</param>
    /// <param name="alignment">Text alignment.</param>
    /// <param name="verticalAlignmentStr">Vertical alignment.</param>
    /// <param name="paddingTop">Top padding.</param>
    /// <param name="paddingBottom">Bottom padding.</param>
    /// <param name="paddingLeft">Left padding.</param>
    /// <param name="paddingRight">Right padding.</param>
    private static void ApplyCellFormatting(Cell cell, string? backgroundColor, string? alignment,
        string? verticalAlignmentStr, double? paddingTop, double? paddingBottom, double? paddingLeft,
        double? paddingRight)
    {
        var cellFormat = cell.CellFormat;

        if (!string.IsNullOrEmpty(backgroundColor))
            cellFormat.Shading.BackgroundPatternColor = ColorHelper.ParseColor(backgroundColor, true);

        if (!string.IsNullOrEmpty(alignment))
        {
            var paragraphs = cell.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
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
    }

    /// <summary>
    ///     Applies text formatting to runs in a cell.
    /// </summary>
    /// <param name="cell">The cell.</param>
    /// <param name="fontName">Font name.</param>
    /// <param name="fontNameAscii">Font name for ASCII.</param>
    /// <param name="fontNameFarEast">Font name for Far East.</param>
    /// <param name="fontSize">Font size.</param>
    /// <param name="bold">Bold setting.</param>
    /// <param name="italic">Italic setting.</param>
    /// <param name="color">Text color.</param>
    private static void ApplyTextFormatting(Cell cell, string? fontName, string? fontNameAscii, string? fontNameFarEast,
        double? fontSize, bool? bold, bool? italic, string? color)
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

    /// <summary>
    ///     Gets the target description for the result message.
    /// </summary>
    /// <param name="applyToTable">Whether applying to entire table.</param>
    /// <param name="applyToRow">Whether applying to entire row.</param>
    /// <param name="applyToColumn">Whether applying to entire column.</param>
    /// <param name="rowIndex">Row index.</param>
    /// <param name="columnIndex">Column index.</param>
    /// <returns>Target description string.</returns>
    private static string GetTargetDescription(bool applyToTable, bool applyToRow, bool applyToColumn, int? rowIndex,
        int? columnIndex)
    {
        return applyToTable
            ? "entire table"
            : applyToRow
                ? $"row {rowIndex}"
                : applyToColumn
                    ? $"column {columnIndex}"
                    : $"cell [{rowIndex}, {columnIndex}]";
    }
}
