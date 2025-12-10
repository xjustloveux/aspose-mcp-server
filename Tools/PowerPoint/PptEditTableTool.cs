using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptEditTableTool : IAsposeTool
{
    public string Description => "Edit table content and format on a PowerPoint slide";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Presentation file path"
            },
            slideIndex = new
            {
                type = "number",
                description = "Slide index (0-based)"
            },
            shapeIndex = new
            {
                type = "number",
                description = "Shape index of the table (0-based)"
            },
            data = new
            {
                type = "array",
                description = "2D array of cell data (optional)",
                items = new
                {
                    type = "array",
                    items = new { type = "string" }
                }
            },
            rowIndex = new
            {
                type = "number",
                description = "Row index to update (optional, if not provided, updates entire table)"
            },
            columnIndex = new
            {
                type = "number",
                description = "Column index to update (optional, if not provided, updates entire table)"
            },
            cellValue = new
            {
                type = "string",
                description = "Cell value (required if rowIndex and columnIndex are provided)"
            }
        },
        required = new[] { "path", "slideIndex", "shapeIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required");
        var dataArray = arguments?["data"]?.AsArray();
        var rowIndex = arguments?["rowIndex"]?.GetValue<int?>();
        var columnIndex = arguments?["columnIndex"]?.GetValue<int?>();
        var cellValue = arguments?["cellValue"]?.GetValue<string>();

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];
        if (shapeIndex < 0 || shapeIndex >= slide.Shapes.Count)
        {
            throw new ArgumentException($"shapeIndex must be between 0 and {slide.Shapes.Count - 1}");
        }

        var shape = slide.Shapes[shapeIndex];
        if (shape is not ITable table)
        {
            throw new ArgumentException($"Shape at index {shapeIndex} is not a table");
        }

        if (rowIndex.HasValue && columnIndex.HasValue)
        {
            // Update single cell
            if (rowIndex.Value < 0 || rowIndex.Value >= table.Rows.Count)
            {
                throw new ArgumentException($"rowIndex must be between 0 and {table.Rows.Count - 1}");
            }
            if (columnIndex.Value < 0 || columnIndex.Value >= table.Columns.Count)
            {
                throw new ArgumentException($"columnIndex must be between 0 and {table.Columns.Count - 1}");
            }
            if (string.IsNullOrEmpty(cellValue))
            {
                throw new ArgumentException("cellValue is required when updating a single cell");
            }
            table[columnIndex.Value, rowIndex.Value].TextFrame.Text = cellValue;
        }
        else if (dataArray != null)
        {
            // Update entire table or range
            for (int i = 0; i < Math.Min(table.Rows.Count, dataArray.Count); i++)
            {
                var rowArray = dataArray[i]?.AsArray();
                if (rowArray != null)
                {
                    for (int j = 0; j < Math.Min(table.Columns.Count, rowArray.Count); j++)
                    {
                        table[j, i].TextFrame.Text = rowArray[j]?.GetValue<string>() ?? "";
                    }
                }
            }
        }
        else
        {
            throw new ArgumentException("Either data array or rowIndex/columnIndex/cellValue must be provided");
        }

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"Table updated on slide {slideIndex}, shape {shapeIndex}");
    }
}

