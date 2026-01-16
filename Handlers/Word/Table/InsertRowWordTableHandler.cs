using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Tables;
using AsposeMcpServer.Core.Handlers;
using WordParagraph = Aspose.Words.Paragraph;

namespace AsposeMcpServer.Handlers.Word.Table;

/// <summary>
///     Handler for inserting rows into Word document tables.
/// </summary>
public class InsertRowWordTableHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "insert_row";

    /// <summary>
    ///     Inserts a row into a table.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: rowIndex.
    ///     Optional: tableIndex (default 0), rowData, insertBefore, sectionIndex.
    /// </param>
    /// <returns>Success message with inserted row index.</returns>
    /// <exception cref="ArgumentException">Thrown when rowIndex is missing or indices are out of range.</exception>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var p = ExtractInsertRowParameters(parameters);

        var doc = context.Document;
        var actualSectionIndex = p.SectionIndex ?? 0;
        if (actualSectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"Section index {actualSectionIndex} out of range");

        var section = doc.Sections[actualSectionIndex];
        var tables = section.Body.GetChildNodes(NodeType.Table, true).Cast<Aspose.Words.Tables.Table>().ToList();
        if (p.TableIndex < 0 || p.TableIndex >= tables.Count)
            throw new ArgumentException($"Table index {p.TableIndex} out of range");

        var table = tables[p.TableIndex];
        if (p.RowIndex < 0 || p.RowIndex >= table.Rows.Count)
            throw new ArgumentException($"Row index {p.RowIndex} out of range");

        JsonArray? dataArray = null;
        if (!string.IsNullOrEmpty(p.RowData))
            try
            {
                dataArray = JsonNode.Parse(p.RowData)?.AsArray();
            }
            catch
            {
                throw new ArgumentException("Invalid rowData JSON format");
            }

        var targetRow = table.Rows[p.RowIndex];
        var newRow = CreateNewRow(doc, targetRow, dataArray);

        if (p.InsertBefore)
            table.InsertBefore(newRow, targetRow);
        else
            table.InsertAfter(newRow, targetRow);

        MarkModified(context);

        var insertedIndex = p.InsertBefore ? p.RowIndex : p.RowIndex + 1;
        return Success($"Successfully inserted row at index {insertedIndex}.");
    }

    /// <summary>
    ///     Creates a new row based on an existing row template.
    /// </summary>
    /// <param name="doc">The document.</param>
    /// <param name="templateRow">The template row to copy format from.</param>
    /// <param name="dataArray">Optional data array for cell content.</param>
    /// <returns>The created row.</returns>
    private static Row CreateNewRow(Document doc, Row templateRow, JsonArray? dataArray)
    {
        var newRow = new Row(doc);

        for (var i = 0; i < templateRow.Cells.Count; i++)
        {
            var sourceCell = templateRow.Cells[i];
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
                    var para = new WordParagraph(doc);
                    AddTextToParagraph(doc, para, cellText);
                    newCell.AppendChild(para);
                }
            }
        }

        return newRow;
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

    private static InsertRowParameters ExtractInsertRowParameters(OperationParameters parameters)
    {
        var rowIndex = parameters.GetOptional<int?>("rowIndex");
        if (!rowIndex.HasValue)
            throw new ArgumentException("rowIndex is required for insert_row operation");

        var tableIndex = parameters.GetOptional("tableIndex", 0);
        var rowData = parameters.GetOptional<string?>("rowData");
        var insertBefore = parameters.GetOptional("insertBefore", false);
        var sectionIndex = parameters.GetOptional<int?>("sectionIndex");

        return new InsertRowParameters(rowIndex.Value, tableIndex, rowData, insertBefore, sectionIndex);
    }

    private sealed record InsertRowParameters(
        int RowIndex,
        int TableIndex,
        string? RowData,
        bool InsertBefore,
        int? SectionIndex);
}
