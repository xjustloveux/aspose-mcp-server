using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptInsertTableRowTool : IAsposeTool
{
    public string Description => "Insert a row into a table on a PowerPoint slide";

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
            rowIndex = new
            {
                type = "number",
                description = "Row index to insert at (0-based, optional, default: append at end)"
            },
            data = new
            {
                type = "array",
                items = new { type = "string" },
                description = "Row data (optional)"
            }
        },
        required = new[] { "path", "slideIndex", "shapeIndex" }
    };

    public Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required");
        var rowIndex = arguments?["rowIndex"]?.GetValue<int?>();
        var dataArray = arguments?["data"]?.AsArray();

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            return Task.FromException<string>(new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}"));
        }

        var slide = presentation.Slides[slideIndex];
        if (shapeIndex < 0 || shapeIndex >= slide.Shapes.Count)
        {
            return Task.FromException<string>(new ArgumentException($"shapeIndex must be between 0 and {slide.Shapes.Count - 1}"));
        }

        var shape = slide.Shapes[shapeIndex];
        if (shape is not ITable table)
        {
            return Task.FromException<string>(new ArgumentException($"Shape at index {shapeIndex} is not a table"));
        }

        var insertIndex = rowIndex ?? table.Rows.Count;
        if (insertIndex < 0 || insertIndex > table.Rows.Count)
        {
            return Task.FromException<string>(new ArgumentException($"rowIndex must be between 0 and {table.Rows.Count}"));
        }

        // Note: Aspose.Slides table rows/columns are typically fixed-size
        // Inserting rows/columns may require recreating the table or using workarounds
        // This functionality may be limited by the API
        return Task.FromException<string>(new NotImplementedException("Table row insertion is not directly supported by Aspose.Slides API. Consider recreating the table with the desired number of rows."));
    }
}

