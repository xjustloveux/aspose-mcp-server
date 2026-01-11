using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Tables;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Table;

/// <summary>
///     Handler for inserting columns into Word document tables.
/// </summary>
public class InsertColumnWordTableHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "insert_column";

    /// <summary>
    ///     Inserts a column into a table.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: columnIndex.
    ///     Optional: tableIndex (default 0), columnData, insertBefore, sectionIndex.
    /// </param>
    /// <returns>Success message with inserted column index.</returns>
    /// <exception cref="ArgumentException">Thrown when columnIndex is missing or indices are out of range.</exception>
    /// <exception cref="InvalidOperationException">Thrown when the table has no rows.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var columnIndex = parameters.GetOptional<int?>("columnIndex");
        if (!columnIndex.HasValue)
            throw new ArgumentException("columnIndex is required for insert_column operation");

        var tableIndex = parameters.GetOptional("tableIndex", 0);
        var columnData = parameters.GetOptional<string?>("columnData");
        var insertBefore = parameters.GetOptional("insertBefore", false);
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
            var newCell = CreateNewCell(doc, row, columnIndex.Value);

            PopulateCellContent(doc, newCell, dataArray, rowIdx);

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

        MarkModified(context);

        var insertedIndex = insertBefore ? columnIndex.Value : columnIndex.Value + 1;
        return Success($"Successfully inserted column at index {insertedIndex}.");
    }

    /// <summary>
    ///     Creates a new cell based on an existing cell template.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="row">The row containing template cells.</param>
    /// <param name="columnIndex">The column index to copy format from.</param>
    /// <returns>The created cell.</returns>
    private static Cell CreateNewCell(Document doc, Row row, int columnIndex)
    {
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

        return newCell;
    }

    /// <summary>
    ///     Populates cell content from data array.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="cell">The cell to populate.</param>
    /// <param name="dataArray">The data array.</param>
    /// <param name="rowIndex">The row index.</param>
    private static void PopulateCellContent(Document doc, Cell cell, JsonArray? dataArray, int rowIndex)
    {
        if (dataArray != null && rowIndex < dataArray.Count)
        {
            var cellDataNode = dataArray[rowIndex];
            var cellText = "";

            if (cellDataNode != null)
                cellText = cellDataNode.GetValueKind() == JsonValueKind.String
                    ? cellDataNode.GetValue<string>()
                    : cellDataNode.ToString();

            if (!string.IsNullOrEmpty(cellText))
            {
                var para = new WordParagraph(doc);
                AddTextToParagraph(doc, para, cellText);
                cell.AppendChild(para);
                return;
            }
        }

        var emptyPara = new WordParagraph(doc);
        cell.AppendChild(emptyPara);
    }

    /// <summary>
    ///     Adds text to a paragraph, handling line breaks.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="para">The paragraph.</param>
    /// <param name="text">The text to add.</param>
    private static void AddTextToParagraph(Document doc, WordParagraph para, string text)
    {
        if (text.Contains('\n'))
        {
            var lines = text.Split('\n');
            for (var lineIdx = 0; lineIdx < lines.Length; lineIdx++)
            {
                if (lineIdx > 0)
                    para.AppendChild(new Run(doc, ControlChar.LineBreak));
                para.AppendChild(new Run(doc, lines[lineIdx]));
            }
        }
        else
        {
            para.AppendChild(new Run(doc, text));
        }
    }
}
