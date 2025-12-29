using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint tables (add, edit, delete, get content, insert row/column, delete row/column)
///     Merges: PptAddTableTool, PptEditTableTool, PptDeleteTableTool, PptGetTableContentTool,
///     PptInsertTableRowTool, PptInsertTableColumnTool, PptDeleteTableRowTool, PptDeleteTableColumnTool,
///     PptEditTableCellTool
/// </summary>
public class PptTableTool : IAsposeTool
{
    public string Description =>
        @"Manage PowerPoint tables. Supports 9 operations: add, edit, delete, get_content, insert_row, insert_column, delete_row, delete_column, edit_cell.

Coordinate unit: 1 inch = 72 points.

Usage examples:
- Add table: ppt_table(operation='add', path='presentation.pptx', slideIndex=0, rows=3, columns=3, x=100, y=100)
- Edit table: ppt_table(operation='edit', path='presentation.pptx', slideIndex=0, shapeIndex=0, data=[['A','B'],['C','D']])
- Delete table: ppt_table(operation='delete', path='presentation.pptx', slideIndex=0, shapeIndex=0)
- Get content: ppt_table(operation='get_content', path='presentation.pptx', slideIndex=0, shapeIndex=0) → Returns JSON with rows, columns, and cell data
- Insert row: ppt_table(operation='insert_row', path='presentation.pptx', slideIndex=0, shapeIndex=0, rowIndex=1)
- Insert column: ppt_table(operation='insert_column', path='presentation.pptx', slideIndex=0, shapeIndex=0, columnIndex=1)
- Delete row: ppt_table(operation='delete_row', path='presentation.pptx', slideIndex=0, shapeIndex=0, rowIndex=1)
- Edit cell: ppt_table(operation='edit_cell', path='presentation.pptx', slideIndex=0, shapeIndex=0, rowIndex=0, columnIndex=0, text='New Value')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add a table (required params: path, slideIndex, rows, columns)
- 'edit': Edit table data (required params: path, slideIndex, shapeIndex, data)
- 'delete': Delete a table (required params: path, slideIndex, shapeIndex)
- 'get_content': Get table content (required params: path, slideIndex, shapeIndex)
- 'insert_row': Insert a row (required params: path, slideIndex, shapeIndex, rowIndex)
- 'insert_column': Insert a column (required params: path, slideIndex, shapeIndex, columnIndex)
- 'delete_row': Delete a row (required params: path, slideIndex, shapeIndex, rowIndex)
- 'delete_column': Delete a column (required params: path, slideIndex, shapeIndex, columnIndex)
- 'edit_cell': Edit cell content (required params: path, slideIndex, shapeIndex, rowIndex, columnIndex, text)",
                @enum = new[]
                {
                    "add", "edit", "delete", "get_content", "insert_row", "insert_column", "delete_row",
                    "delete_column", "edit_cell"
                }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path (required for all operations)"
            },
            slideIndex = new
            {
                type = "number",
                description = "Slide index (0-based)"
            },
            shapeIndex = new
            {
                type = "number",
                description = "Shape index of the table (0-based, required for most operations)"
            },
            rows = new
            {
                type = "number",
                description = "Number of rows (required for add)"
            },
            columns = new
            {
                type = "number",
                description = "Number of columns (required for add)"
            },
            x = new
            {
                type = "number",
                description = "X position in points (optional for add, defaults to 50)"
            },
            y = new
            {
                type = "number",
                description = "Y position in points (optional for add, defaults to 50)"
            },
            data = new
            {
                type = "array",
                description = "2D array of cell data (optional, for add/edit)",
                items = new
                {
                    type = "array",
                    items = new { type = "string" }
                }
            },
            rowIndex = new
            {
                type = "number",
                description = "Row index (0-based, required for insert_row/delete_row/edit_cell)"
            },
            columnIndex = new
            {
                type = "number",
                description = "Column index (0-based, required for insert_column/delete_column/edit_cell)"
            },
            text = new
            {
                type = "string",
                description = "Cell text content (required for edit_cell)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for add/edit/delete operations, defaults to input path)"
            }
        },
        required = new[] { "operation", "path", "slideIndex" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments.
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters.</param>
    /// <returns>Result message as a string.</returns>
    /// <exception cref="ArgumentException">Thrown when operation is unknown.</exception>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");

        return operation.ToLower() switch
        {
            "add" => await AddTableAsync(path, outputPath, slideIndex, arguments),
            "edit" => await EditTableAsync(path, outputPath, slideIndex, arguments),
            "delete" => await DeleteTableAsync(path, outputPath, slideIndex, arguments),
            "get_content" => await GetTableContentAsync(path, slideIndex, arguments),
            "insert_row" => await InsertRowAsync(path, outputPath, slideIndex, arguments),
            "insert_column" => await InsertColumnAsync(path, outputPath, slideIndex, arguments),
            "delete_row" => await DeleteRowAsync(path, outputPath, slideIndex, arguments),
            "delete_column" => await DeleteColumnAsync(path, outputPath, slideIndex, arguments),
            "edit_cell" => await EditCellAsync(path, outputPath, slideIndex, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a table to a slide.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="slideIndex">Slide index (0-based).</param>
    /// <param name="arguments">JSON arguments containing rows, columns, x, y, optional data.</param>
    /// <returns>Success message with table dimensions.</returns>
    /// <exception cref="ArgumentException">Thrown when rows or columns are out of valid range (1-1000).</exception>
    private Task<string> AddTableAsync(string path, string outputPath, int slideIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var rows = ArgumentHelper.GetInt(arguments, "rows");
            var columns = ArgumentHelper.GetInt(arguments, "columns");
            var x = ArgumentHelper.GetFloat(arguments, "x", 50);
            var y = ArgumentHelper.GetFloat(arguments, "y", 50);
            var dataArray = ArgumentHelper.GetArray(arguments, "data", false);

            if (rows <= 0 || rows > 1000) throw new ArgumentException("rows must be between 1 and 1000");
            if (columns <= 0 || columns > 1000) throw new ArgumentException("columns must be between 1 and 1000");

            using var presentation = new Presentation(path);
            if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
                throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");

            var slide = presentation.Slides[slideIndex];

            var columnWidths = new double[columns];
            var rowHeights = new double[rows];

            for (var i = 0; i < columns; i++)
                columnWidths[i] = 100;
            for (var i = 0; i < rows; i++)
                rowHeights[i] = 50;

            var table = slide.Shapes.AddTable(x, y, columnWidths, rowHeights);

            if (dataArray != null)
                for (var i = 0; i < Math.Min(rows, dataArray.Count); i++)
                {
                    var rowArray = dataArray[i]?.AsArray();
                    if (rowArray != null)
                        for (var j = 0; j < Math.Min(columns, rowArray.Count); j++)
                            table[i, j].TextFrame.Text = rowArray[j]?.GetValue<string>() ?? "";
                }

            presentation.Save(outputPath, SaveFormat.Pptx);

            return $"Table ({rows}x{columns}) added to slide {slideIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits table data.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="slideIndex">Slide index (0-based).</param>
    /// <param name="arguments">JSON arguments containing shapeIndex, optional data.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when shape is not a table.</exception>
    private Task<string> EditTableAsync(string path, string outputPath, int slideIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
            var dataArray = ArgumentHelper.GetArray(arguments, "data", false);

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var shape = PowerPointHelper.GetShape(slide, shapeIndex);
            if (shape is not ITable table) throw new ArgumentException($"Shape at index {shapeIndex} is not a table");

            if (dataArray != null)
                for (var i = 0; i < Math.Min(table.Rows.Count, dataArray.Count); i++)
                {
                    var rowArray = dataArray[i]?.AsArray();
                    if (rowArray != null)
                        for (var j = 0; j < Math.Min(table.Columns.Count, rowArray.Count); j++)
                            table[i, j].TextFrame.Text = rowArray[j]?.GetValue<string>() ?? "";
                }

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Table on slide {slideIndex}, shape {shapeIndex} updated. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes a table from a slide.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="slideIndex">Slide index (0-based).</param>
    /// <param name="arguments">JSON arguments containing shapeIndex.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when shape is not a table.</exception>
    private Task<string> DeleteTableAsync(string path, string outputPath, int slideIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var shape = PowerPointHelper.GetShape(slide, shapeIndex);
            if (shape is not ITable) throw new ArgumentException($"Shape at index {shapeIndex} is not a table");

            slide.Shapes.Remove(shape);

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Table on slide {slideIndex}, shape {shapeIndex} deleted. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets table content as JSON.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="slideIndex">Slide index (0-based).</param>
    /// <param name="arguments">JSON arguments containing shapeIndex.</param>
    /// <returns>JSON string with slideIndex, shapeIndex, columns, rows, and data array.</returns>
    /// <exception cref="ArgumentException">Thrown when shape is not a table.</exception>
    private Task<string> GetTableContentAsync(string path, int slideIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var shape = PowerPointHelper.GetShape(slide, shapeIndex);
            if (shape is not ITable table) throw new ArgumentException($"Shape at index {shapeIndex} is not a table");

            var rowsList = new List<object>();
            for (var i = 0; i < table.Rows.Count; i++)
            {
                var cellsList = new List<object>();
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
                shapeIndex,
                columns = table.Columns.Count,
                rows = table.Rows.Count,
                data = rowsList
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }

    /// <summary>
    ///     Inserts a row into a table by cloning an existing row.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="slideIndex">Slide index (0-based).</param>
    /// <param name="arguments">JSON arguments containing shapeIndex, rowIndex.</param>
    /// <returns>Success message with new row count.</returns>
    /// <exception cref="ArgumentException">Thrown when shape is not a table or rowIndex is out of range.</exception>
    private Task<string> InsertRowAsync(string path, string outputPath, int slideIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
            var rowIndex = ArgumentHelper.GetInt(arguments, "rowIndex");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var shape = PowerPointHelper.GetShape(slide, shapeIndex);
            if (shape is not ITable table) throw new ArgumentException("Shape is not a table");

            if (rowIndex < 0 || rowIndex > table.Rows.Count)
                throw new ArgumentException($"rowIndex {rowIndex} is out of range (0-{table.Rows.Count})");

            var templateRow = table.Rows[Math.Min(rowIndex, table.Rows.Count - 1)];
            table.Rows.InsertClone(rowIndex, templateRow, false);

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Row inserted at index {rowIndex}. Table now has {table.Rows.Count} rows. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Inserts a column into a table by cloning an existing column.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="slideIndex">Slide index (0-based).</param>
    /// <param name="arguments">JSON arguments containing shapeIndex, columnIndex.</param>
    /// <returns>Success message with new column count.</returns>
    /// <exception cref="ArgumentException">Thrown when shape is not a table or columnIndex is out of range.</exception>
    private Task<string> InsertColumnAsync(string path, string outputPath, int slideIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
            var columnIndex = ArgumentHelper.GetInt(arguments, "columnIndex");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var shape = PowerPointHelper.GetShape(slide, shapeIndex);
            if (shape is not ITable table) throw new ArgumentException("Shape is not a table");

            if (columnIndex < 0 || columnIndex > table.Columns.Count)
                throw new ArgumentException($"columnIndex {columnIndex} is out of range (0-{table.Columns.Count})");

            var templateColumn = table.Columns[Math.Min(columnIndex, table.Columns.Count - 1)];
            table.Columns.InsertClone(columnIndex, templateColumn, false);

            presentation.Save(outputPath, SaveFormat.Pptx);
            return
                $"Column inserted at index {columnIndex}. Table now has {table.Columns.Count} columns. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes a row from a table.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="slideIndex">Slide index (0-based).</param>
    /// <param name="arguments">JSON arguments containing shapeIndex, rowIndex.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when shape is not a table.</exception>
    private Task<string> DeleteRowAsync(string path, string outputPath, int slideIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
            var rowIndex = ArgumentHelper.GetInt(arguments, "rowIndex");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var shape = PowerPointHelper.GetShape(slide, shapeIndex);
            if (shape is not ITable table) throw new ArgumentException("Shape is not a table");

            table.Rows.RemoveAt(rowIndex, false);
            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Row {rowIndex} deleted. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes a column from a table.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="slideIndex">Slide index (0-based).</param>
    /// <param name="arguments">JSON arguments containing shapeIndex, columnIndex.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when shape is not a table.</exception>
    private Task<string> DeleteColumnAsync(string path, string outputPath, int slideIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
            var columnIndex = ArgumentHelper.GetInt(arguments, "columnIndex");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var shape = PowerPointHelper.GetShape(slide, shapeIndex);
            if (shape is not ITable table) throw new ArgumentException("Shape is not a table");

            table.Columns.RemoveAt(columnIndex, false);
            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Column {columnIndex} deleted. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits a cell in a table.
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="slideIndex">Slide index (0-based).</param>
    /// <param name="arguments">JSON arguments containing shapeIndex, rowIndex, columnIndex, text.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="ArgumentException">Thrown when shape is not a table or row/column index is out of range.</exception>
    private Task<string> EditCellAsync(string path, string outputPath, int slideIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var shapeIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
            var rowIndex = ArgumentHelper.GetInt(arguments, "rowIndex");
            var columnIndex = ArgumentHelper.GetInt(arguments, "columnIndex");
            var text = ArgumentHelper.GetString(arguments, "text");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var shape = PowerPointHelper.GetShape(slide, shapeIndex);
            if (shape is not ITable table) throw new ArgumentException("Shape is not a table");

            if (rowIndex < 0 || rowIndex >= table.Rows.Count)
                throw new ArgumentException($"rowIndex {rowIndex} is out of range (0-{table.Rows.Count - 1})");
            if (columnIndex < 0 || columnIndex >= table.Columns.Count)
                throw new ArgumentException($"columnIndex {columnIndex} is out of range (0-{table.Columns.Count - 1})");

            table[rowIndex, columnIndex].TextFrame.Text = text;
            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Cell [{rowIndex},{columnIndex}] updated. Output: {outputPath}";
        });
    }
}