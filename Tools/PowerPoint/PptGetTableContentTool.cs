using System.Text.Json.Nodes;
using System.Text;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptGetTableContentTool : IAsposeTool
{
    public string Description => "Get table content and structure from a PowerPoint slide";

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
            }
        },
        required = new[] { "path", "slideIndex", "shapeIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var shapeIndex = arguments?["shapeIndex"]?.GetValue<int>() ?? throw new ArgumentException("shapeIndex is required");

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

        var sb = new StringBuilder();
        sb.AppendLine($"Table: {table.Columns.Count} columns x {table.Rows.Count} rows");
        sb.AppendLine();

        for (int i = 0; i < table.Rows.Count; i++)
        {
            var row = new List<string>();
            for (int j = 0; j < table.Columns.Count; j++)
            {
                var cell = table[j, i];
                var text = cell.TextFrame?.Text ?? "";
                row.Add(text);
            }
            sb.AppendLine($"Row {i}: {string.Join(" | ", row)}");
        }

        return await Task.FromResult(sb.ToString());
    }
}

