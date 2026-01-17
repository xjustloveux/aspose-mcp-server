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
        var p = ExtractEditCellFormatParameters(parameters);

        var doc = context.Document;
        var actualSectionIndex = p.SectionIndex ?? 0;
        if (actualSectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"Section index {actualSectionIndex} out of range");

        var section = doc.Sections[actualSectionIndex];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Aspose.Words.Tables.Table>().ToList();
        if (p.TableIndex < 0 || p.TableIndex >= tables.Count)
            throw new ArgumentException($"Table index {p.TableIndex} out of range");

        var table = tables[p.TableIndex];

        var targetCells =
            GetTargetCells(table, p.RowIndex, p.ColumnIndex, p.ApplyToRow, p.ApplyToColumn, p.ApplyToTable);

        if (targetCells.Count == 0)
            throw new ArgumentException("No target cells found");

        var hasTextFormatting = !string.IsNullOrEmpty(p.FontName) || !string.IsNullOrEmpty(p.FontNameAscii) ||
                                !string.IsNullOrEmpty(p.FontNameFarEast) || p.FontSize.HasValue ||
                                p.Bold.HasValue || p.Italic.HasValue || !string.IsNullOrEmpty(p.Color);

        foreach (var cell in targetCells)
        {
            ApplyCellFormatting(cell, p);

            if (hasTextFormatting)
                ApplyTextFormatting(cell, p);
        }

        var targetDescription =
            GetTargetDescription(p.ApplyToTable, p.ApplyToRow, p.ApplyToColumn, p.RowIndex, p.ColumnIndex);

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
    private static List<Cell> GetTargetCells( // NOSONAR S3776 - Mutually exclusive cell selection
        Aspose.Words.Tables.Table table, int? rowIndex, int? columnIndex,
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
            targetCells.AddRange(table.Rows.Cast<Row>()
                .Where(row => columnIndex.Value < row.Cells.Count)
                .Select(row => row.Cells[columnIndex.Value]));
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
    /// <param name="p">The edit cell format parameters.</param>
    private static void ApplyCellFormatting(Cell cell, EditCellFormatParameters p)
    {
        var cellFormat = cell.CellFormat;

        if (!string.IsNullOrEmpty(p.BackgroundColor))
            cellFormat.Shading.BackgroundPatternColor = ColorHelper.ParseColor(p.BackgroundColor, true);

        if (!string.IsNullOrEmpty(p.Alignment))
        {
            var paragraphs = cell.GetChildNodes(NodeType.Paragraph, true).Cast<WordParagraph>().ToList();
            foreach (var para in paragraphs)
                para.ParagraphFormat.Alignment = p.Alignment.ToLower() switch
                {
                    "center" => ParagraphAlignment.Center,
                    "right" => ParagraphAlignment.Right,
                    "justify" => ParagraphAlignment.Justify,
                    _ => ParagraphAlignment.Left
                };
        }

        if (!string.IsNullOrEmpty(p.VerticalAlignmentStr))
            cellFormat.VerticalAlignment = p.VerticalAlignmentStr.ToLower() switch
            {
                "center" => CellVerticalAlignment.Center,
                "bottom" => CellVerticalAlignment.Bottom,
                _ => CellVerticalAlignment.Top
            };

        if (p.PaddingTop.HasValue)
            cellFormat.TopPadding = p.PaddingTop.Value;
        if (p.PaddingBottom.HasValue)
            cellFormat.BottomPadding = p.PaddingBottom.Value;
        if (p.PaddingLeft.HasValue)
            cellFormat.LeftPadding = p.PaddingLeft.Value;
        if (p.PaddingRight.HasValue)
            cellFormat.RightPadding = p.PaddingRight.Value;
    }

    /// <summary>
    ///     Applies text formatting to runs in a cell.
    /// </summary>
    /// <param name="cell">The cell.</param>
    /// <param name="p">The edit cell format parameters.</param>
    private static void ApplyTextFormatting(Cell cell, EditCellFormatParameters p)
    {
        var runs = cell.GetChildNodes(NodeType.Run, true).Cast<Run>().ToList();
        foreach (var run in runs)
            FontHelper.Word.ApplyFontSettings(
                run,
                p.FontName,
                p.FontNameAscii,
                p.FontNameFarEast,
                p.FontSize,
                p.Bold,
                p.Italic,
                color: p.Color
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
        if (applyToTable) return "entire table";
        if (applyToRow) return $"row {rowIndex}";
        if (applyToColumn) return $"column {columnIndex}";
        return $"cell [{rowIndex}, {columnIndex}]";
    }

    /// <summary>
    ///     Extracts edit cell format parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted edit cell format parameters.</returns>
    private static EditCellFormatParameters ExtractEditCellFormatParameters(OperationParameters parameters)
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

        return new EditCellFormatParameters(
            tableIndex,
            rowIndex,
            columnIndex,
            applyToRow,
            applyToColumn,
            applyToTable,
            backgroundColor,
            alignment,
            verticalAlignmentStr,
            paddingTop,
            paddingBottom,
            paddingLeft,
            paddingRight,
            fontName,
            fontNameAscii,
            fontNameFarEast,
            fontSize,
            bold,
            italic,
            color,
            sectionIndex);
    }

    /// <summary>
    ///     Record to hold edit cell format parameters.
    /// </summary>
    /// <param name="TableIndex">The table index.</param>
    /// <param name="RowIndex">The row index.</param>
    /// <param name="ColumnIndex">The column index.</param>
    /// <param name="ApplyToRow">Whether to apply to entire row.</param>
    /// <param name="ApplyToColumn">Whether to apply to entire column.</param>
    /// <param name="ApplyToTable">Whether to apply to entire table.</param>
    /// <param name="BackgroundColor">The background color.</param>
    /// <param name="Alignment">The horizontal alignment.</param>
    /// <param name="VerticalAlignmentStr">The vertical alignment.</param>
    /// <param name="PaddingTop">The top padding.</param>
    /// <param name="PaddingBottom">The bottom padding.</param>
    /// <param name="PaddingLeft">The left padding.</param>
    /// <param name="PaddingRight">The right padding.</param>
    /// <param name="FontName">The font name.</param>
    /// <param name="FontNameAscii">The ASCII font name.</param>
    /// <param name="FontNameFarEast">The Far East font name.</param>
    /// <param name="FontSize">The font size.</param>
    /// <param name="Bold">Whether to apply bold.</param>
    /// <param name="Italic">Whether to apply italic.</param>
    /// <param name="Color">The font color.</param>
    /// <param name="SectionIndex">The section index.</param>
    private sealed record EditCellFormatParameters(
        int TableIndex,
        int? RowIndex,
        int? ColumnIndex,
        bool ApplyToRow,
        bool ApplyToColumn,
        bool ApplyToTable,
        string? BackgroundColor,
        string? Alignment,
        string? VerticalAlignmentStr,
        double? PaddingTop,
        double? PaddingBottom,
        double? PaddingLeft,
        double? PaddingRight,
        string? FontName,
        string? FontNameAscii,
        string? FontNameFarEast,
        double? FontSize,
        bool? Bold,
        bool? Italic,
        string? Color,
        int? SectionIndex);
}
