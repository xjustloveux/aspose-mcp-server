using System.ComponentModel;
using System.Text.Json;
using Aspose.Slides;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint tables (add, edit, delete, get content, insert row/column, delete row/column)
///     Merges: PptAddTableTool, PptEditTableTool, PptDeleteTableTool, PptGetTableContentTool,
///     PptInsertTableRowTool, PptInsertTableColumnTool, PptDeleteTableRowTool, PptDeleteTableColumnTool,
///     PptEditTableCellTool
/// </summary>
[McpServerToolType]
public class PptTableTool
{
    /// <summary>
    ///     Identity accessor for session isolation.
    /// </summary>
    private readonly ISessionIdentityAccessor? _identityAccessor;

    /// <summary>
    ///     Session manager for document session handling.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptTableTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory editing.</param>
    /// <param name="identityAccessor">Optional identity accessor for session isolation.</param>
    public PptTableTool(DocumentSessionManager? sessionManager = null,
        ISessionIdentityAccessor? identityAccessor = null)
    {
        _sessionManager = sessionManager;
        _identityAccessor = identityAccessor;
    }

    /// <summary>
    ///     Executes a PowerPoint table operation (add, edit, delete, get_content, insert_row, insert_column, delete_row,
    ///     delete_column, edit_cell).
    /// </summary>
    /// <param name="operation">
    ///     The operation to perform: add, edit, delete, get_content, insert_row, insert_column,
    ///     delete_row, delete_column, edit_cell.
    /// </param>
    /// <param name="path">Presentation file path (required if no sessionId).</param>
    /// <param name="sessionId">Session ID for in-memory editing.</param>
    /// <param name="outputPath">Output file path (optional, defaults to input path).</param>
    /// <param name="slideIndex">Slide index (0-based).</param>
    /// <param name="shapeIndex">Shape index of the table (0-based, required for most operations).</param>
    /// <param name="rows">Number of rows (required for add).</param>
    /// <param name="columns">Number of columns (required for add).</param>
    /// <param name="x">X position in points (optional for add, defaults to 50).</param>
    /// <param name="y">Y position in points (optional for add, defaults to 50).</param>
    /// <param name="data">2D array of cell data as JSON (optional, for add/edit).</param>
    /// <param name="rowIndex">Row index (0-based, required for insert_row/delete_row/edit_cell).</param>
    /// <param name="columnIndex">Column index (0-based, required for insert_column/delete_column/edit_cell).</param>
    /// <param name="text">Cell text content (required for edit_cell).</param>
    /// <returns>A message indicating the result of the operation, or JSON data for get operations.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or the operation is unknown.</exception>
    [McpServerTool(Name = "ppt_table")]
    [Description(
        @"Manage PowerPoint tables. Supports 9 operations: add, edit, delete, get_content, insert_row, insert_column, delete_row, delete_column, edit_cell.

Coordinate unit: 1 inch = 72 points.

Usage examples:
- Add table: ppt_table(operation='add', path='presentation.pptx', slideIndex=0, rows=3, columns=3, x=100, y=100)
- Edit table: ppt_table(operation='edit', path='presentation.pptx', slideIndex=0, shapeIndex=0, data='[[""A"",""B""],[""C"",""D""]]')
- Delete table: ppt_table(operation='delete', path='presentation.pptx', slideIndex=0, shapeIndex=0)
- Get content: ppt_table(operation='get_content', path='presentation.pptx', slideIndex=0, shapeIndex=0)
- Insert row: ppt_table(operation='insert_row', path='presentation.pptx', slideIndex=0, shapeIndex=0, rowIndex=1)
- Insert column: ppt_table(operation='insert_column', path='presentation.pptx', slideIndex=0, shapeIndex=0, columnIndex=1)
- Delete row: ppt_table(operation='delete_row', path='presentation.pptx', slideIndex=0, shapeIndex=0, rowIndex=1)
- Edit cell: ppt_table(operation='edit_cell', path='presentation.pptx', slideIndex=0, shapeIndex=0, rowIndex=0, columnIndex=0, text='New Value')")]
    public string Execute(
        [Description(@"Operation to perform.
- 'add': Add a table (required params: path, slideIndex, rows, columns)
- 'edit': Edit table data (required params: path, slideIndex, shapeIndex, data)
- 'delete': Delete a table (required params: path, slideIndex, shapeIndex)
- 'get_content': Get table content (required params: path, slideIndex, shapeIndex)
- 'insert_row': Insert a row (required params: path, slideIndex, shapeIndex, rowIndex)
- 'insert_column': Insert a column (required params: path, slideIndex, shapeIndex, columnIndex)
- 'delete_row': Delete a row (required params: path, slideIndex, shapeIndex, rowIndex)
- 'delete_column': Delete a column (required params: path, slideIndex, shapeIndex, columnIndex)
- 'edit_cell': Edit cell content (required params: path, slideIndex, shapeIndex, rowIndex, columnIndex, text)")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (optional, defaults to input path)")]
        string? outputPath = null,
        [Description("Slide index (0-based)")] int slideIndex = 0,
        [Description("Shape index of the table (0-based, required for most operations)")]
        int? shapeIndex = null,
        [Description("Number of rows (required for add)")]
        int? rows = null,
        [Description("Number of columns (required for add)")]
        int? columns = null,
        [Description("X position in points (optional for add, defaults to 50)")]
        float x = 50,
        [Description("Y position in points (optional for add, defaults to 50)")]
        float y = 50,
        [Description("2D array of cell data as JSON (optional, for add/edit)")]
        string? data = null,
        [Description("Row index (0-based, required for insert_row/delete_row/edit_cell)")]
        int? rowIndex = null,
        [Description("Column index (0-based, required for insert_column/delete_column/edit_cell)")]
        int? columnIndex = null,
        [Description("Cell text content (required for edit_cell)")]
        string? text = null)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path, _identityAccessor);

        return operation.ToLower() switch
        {
            "add" => AddTable(ctx, outputPath, slideIndex, rows, columns, x, y, data),
            "edit" => EditTable(ctx, outputPath, slideIndex, shapeIndex, data),
            "delete" => DeleteTable(ctx, outputPath, slideIndex, shapeIndex),
            "get_content" => GetTableContent(ctx, slideIndex, shapeIndex),
            "insert_row" => InsertRow(ctx, outputPath, slideIndex, shapeIndex, rowIndex),
            "insert_column" => InsertColumn(ctx, outputPath, slideIndex, shapeIndex, columnIndex),
            "delete_row" => DeleteRow(ctx, outputPath, slideIndex, shapeIndex, rowIndex),
            "delete_column" => DeleteColumn(ctx, outputPath, slideIndex, shapeIndex, columnIndex),
            "edit_cell" => EditCell(ctx, outputPath, slideIndex, shapeIndex, rowIndex, columnIndex, text),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a table to a slide.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="rows">The number of rows.</param>
    /// <param name="columns">The number of columns.</param>
    /// <param name="x">The X position in points.</param>
    /// <param name="y">The Y position in points.</param>
    /// <param name="dataJson">The table data as JSON 2D array.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when rows or columns are not provided or invalid.</exception>
    private static string AddTable(DocumentContext<Presentation> ctx, string? outputPath, int slideIndex,
        int? rows, int? columns, float x, float y, string? dataJson)
    {
        if (!rows.HasValue)
            throw new ArgumentException("rows is required for add operation");
        if (!columns.HasValue)
            throw new ArgumentException("columns is required for add operation");
        if (rows.Value <= 0 || rows.Value > 1000)
            throw new ArgumentException("rows must be between 1 and 1000");
        if (columns.Value <= 0 || columns.Value > 1000)
            throw new ArgumentException("columns must be between 1 and 1000");

        var presentation = ctx.Document;
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");

        var slide = presentation.Slides[slideIndex];

        var columnWidths = new double[columns.Value];
        var rowHeights = new double[rows.Value];

        for (var i = 0; i < columns.Value; i++)
            columnWidths[i] = 100;
        for (var i = 0; i < rows.Value; i++)
            rowHeights[i] = 50;

        var table = slide.Shapes.AddTable(x, y, columnWidths, rowHeights);

        if (!string.IsNullOrWhiteSpace(dataJson))
        {
            var dataArray = JsonSerializer.Deserialize<string[][]>(dataJson);
            if (dataArray != null)
                for (var i = 0; i < Math.Min(rows.Value, dataArray.Length); i++)
                {
                    var rowArray = dataArray[i];
                    for (var j = 0; j < Math.Min(columns.Value, rowArray.Length); j++)
                        table[i, j].TextFrame.Text = rowArray[j];
                }
        }

        ctx.Save(outputPath);

        return $"Table ({rows.Value}x{columns.Value}) added to slide {slideIndex}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Edits table data.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index of the table (0-based).</param>
    /// <param name="dataJson">The table data as JSON 2D array.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when shapeIndex is not provided or shape is not a table.</exception>
    private static string EditTable(DocumentContext<Presentation> ctx, string? outputPath, int slideIndex,
        int? shapeIndex, string? dataJson)
    {
        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for edit operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex.Value);
        if (shape is not ITable table) throw new ArgumentException($"Shape at index {shapeIndex.Value} is not a table");

        if (!string.IsNullOrWhiteSpace(dataJson))
        {
            var dataArray = JsonSerializer.Deserialize<string[][]>(dataJson);
            if (dataArray != null)
                for (var i = 0; i < Math.Min(table.Rows.Count, dataArray.Length); i++)
                {
                    var rowArray = dataArray[i];
                    for (var j = 0; j < Math.Min(table.Columns.Count, rowArray.Length); j++)
                        table[i, j].TextFrame.Text = rowArray[j];
                }
        }

        ctx.Save(outputPath);
        return $"Table on slide {slideIndex}, shape {shapeIndex.Value} updated. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Deletes a table from a slide.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index of the table (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when shapeIndex is not provided or shape is not a table.</exception>
    private static string DeleteTable(DocumentContext<Presentation> ctx, string? outputPath, int slideIndex,
        int? shapeIndex)
    {
        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for delete operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex.Value);
        if (shape is not ITable) throw new ArgumentException($"Shape at index {shapeIndex.Value} is not a table");

        slide.Shapes.Remove(shape);

        ctx.Save(outputPath);
        return $"Table on slide {slideIndex}, shape {shapeIndex.Value} deleted. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Gets table content as JSON.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index of the table (0-based).</param>
    /// <returns>A JSON string containing the table content.</returns>
    /// <exception cref="ArgumentException">Thrown when shapeIndex is not provided or shape is not a table.</exception>
    private static string GetTableContent(DocumentContext<Presentation> ctx, int slideIndex, int? shapeIndex)
    {
        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for get_content operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex.Value);
        if (shape is not ITable table) throw new ArgumentException($"Shape at index {shapeIndex.Value} is not a table");

        List<object> rowsList = [];
        for (var i = 0; i < table.Rows.Count; i++)
        {
            List<object> cellsList = [];
            for (var j = 0; j < table.Columns.Count; j++)
            {
                var cell = table[i, j];
                var text = cell.TextFrame?.Text ?? "";
                cellsList.Add(new
                {
                    columnIndex = j,
                    text
                });
            }

            rowsList.Add(new
            {
                rowIndex = i,
                cells = cellsList
            });
        }

        var result = new
        {
            slideIndex,
            shapeIndex = shapeIndex.Value,
            columns = table.Columns.Count,
            rows = table.Rows.Count,
            data = rowsList
        };

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Inserts a row into a table by cloning an existing row.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index of the table (0-based).</param>
    /// <param name="rowIndex">The row index where to insert (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when shapeIndex or rowIndex is not provided, shape is not a table, or index
    ///     is out of range.
    /// </exception>
    private static string InsertRow(DocumentContext<Presentation> ctx, string? outputPath, int slideIndex,
        int? shapeIndex, int? rowIndex)
    {
        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for insert_row operation");
        if (!rowIndex.HasValue)
            throw new ArgumentException("rowIndex is required for insert_row operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex.Value);
        if (shape is not ITable table) throw new ArgumentException("Shape is not a table");

        if (rowIndex.Value < 0 || rowIndex.Value > table.Rows.Count)
            throw new ArgumentException($"rowIndex {rowIndex.Value} is out of range (0-{table.Rows.Count})");

        var templateRow = table.Rows[Math.Min(rowIndex.Value, table.Rows.Count - 1)];
        table.Rows.InsertClone(rowIndex.Value, templateRow, false);

        ctx.Save(outputPath);
        return
            $"Row inserted at index {rowIndex.Value}. Table now has {table.Rows.Count} rows. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Inserts a column into a table by cloning an existing column.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index of the table (0-based).</param>
    /// <param name="columnIndex">The column index where to insert (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when shapeIndex or columnIndex is not provided, shape is not a table, or
    ///     index is out of range.
    /// </exception>
    private static string InsertColumn(DocumentContext<Presentation> ctx, string? outputPath, int slideIndex,
        int? shapeIndex, int? columnIndex)
    {
        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for insert_column operation");
        if (!columnIndex.HasValue)
            throw new ArgumentException("columnIndex is required for insert_column operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex.Value);
        if (shape is not ITable table) throw new ArgumentException("Shape is not a table");

        if (columnIndex.Value < 0 || columnIndex.Value > table.Columns.Count)
            throw new ArgumentException($"columnIndex {columnIndex.Value} is out of range (0-{table.Columns.Count})");

        var templateColumn = table.Columns[Math.Min(columnIndex.Value, table.Columns.Count - 1)];
        table.Columns.InsertClone(columnIndex.Value, templateColumn, false);

        ctx.Save(outputPath);
        return
            $"Column inserted at index {columnIndex.Value}. Table now has {table.Columns.Count} columns. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Deletes a row from a table.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index of the table (0-based).</param>
    /// <param name="rowIndex">The row index to delete (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when shapeIndex or rowIndex is not provided, or shape is not a table.</exception>
    private static string DeleteRow(DocumentContext<Presentation> ctx, string? outputPath, int slideIndex,
        int? shapeIndex, int? rowIndex)
    {
        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for delete_row operation");
        if (!rowIndex.HasValue)
            throw new ArgumentException("rowIndex is required for delete_row operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex.Value);
        if (shape is not ITable table) throw new ArgumentException("Shape is not a table");

        table.Rows.RemoveAt(rowIndex.Value, false);
        ctx.Save(outputPath);
        return $"Row {rowIndex.Value} deleted. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Deletes a column from a table.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index of the table (0-based).</param>
    /// <param name="columnIndex">The column index to delete (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when shapeIndex or columnIndex is not provided, or shape is not a table.</exception>
    private static string DeleteColumn(DocumentContext<Presentation> ctx, string? outputPath, int slideIndex,
        int? shapeIndex, int? columnIndex)
    {
        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for delete_column operation");
        if (!columnIndex.HasValue)
            throw new ArgumentException("columnIndex is required for delete_column operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex.Value);
        if (shape is not ITable table) throw new ArgumentException("Shape is not a table");

        table.Columns.RemoveAt(columnIndex.Value, false);
        ctx.Save(outputPath);
        return $"Column {columnIndex.Value} deleted. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Edits a cell in a table.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="shapeIndex">The shape index of the table (0-based).</param>
    /// <param name="rowIndex">The row index of the cell (0-based).</param>
    /// <param name="columnIndex">The column index of the cell (0-based).</param>
    /// <param name="text">The new text content for the cell.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when required parameters are not provided, shape is not a table, or indices
    ///     are out of range.
    /// </exception>
    private static string EditCell(DocumentContext<Presentation> ctx, string? outputPath, int slideIndex,
        int? shapeIndex, int? rowIndex, int? columnIndex, string? text)
    {
        if (!shapeIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for edit_cell operation");
        if (!rowIndex.HasValue)
            throw new ArgumentException("rowIndex is required for edit_cell operation");
        if (!columnIndex.HasValue)
            throw new ArgumentException("columnIndex is required for edit_cell operation");
        if (text == null)
            throw new ArgumentException("text is required for edit_cell operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var shape = PowerPointHelper.GetShape(slide, shapeIndex.Value);
        if (shape is not ITable table) throw new ArgumentException("Shape is not a table");

        if (rowIndex.Value < 0 || rowIndex.Value >= table.Rows.Count)
            throw new ArgumentException($"rowIndex {rowIndex.Value} is out of range (0-{table.Rows.Count - 1})");
        if (columnIndex.Value < 0 || columnIndex.Value >= table.Columns.Count)
            throw new ArgumentException(
                $"columnIndex {columnIndex.Value} is out of range (0-{table.Columns.Count - 1})");

        table[rowIndex.Value, columnIndex.Value].TextFrame.Text = text;
        ctx.Save(outputPath);
        return $"Cell [{rowIndex.Value},{columnIndex.Value}] updated. {ctx.GetOutputMessage(outputPath)}";
    }
}