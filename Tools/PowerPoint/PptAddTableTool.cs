using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptAddTableTool : IAsposeTool
{
    public string Description => "Add a table to a PowerPoint slide";

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
            rows = new
            {
                type = "number",
                description = "Number of rows"
            },
            columns = new
            {
                type = "number",
                description = "Number of columns"
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
            }
        },
        required = new[] { "path", "slideIndex", "rows", "columns" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var slideIndex = arguments?["slideIndex"]?.GetValue<int>() ?? throw new ArgumentException("slideIndex is required");
        var rows = arguments?["rows"]?.GetValue<int>() ?? throw new ArgumentException("rows is required");
        var columns = arguments?["columns"]?.GetValue<int>() ?? throw new ArgumentException("columns is required");
        var dataArray = arguments?["data"]?.AsArray();

        if (rows <= 0 || rows > 1000)
        {
            throw new ArgumentException("rows must be between 1 and 1000");
        }
        if (columns <= 0 || columns > 1000)
        {
            throw new ArgumentException("columns must be between 1 and 1000");
        }

        using var presentation = new Presentation(path);
        if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
        {
            throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");
        }

        var slide = presentation.Slides[slideIndex];

        double[] columnWidths = new double[columns];
        double[] rowHeights = new double[rows];
        
        for (int i = 0; i < columns; i++)
            columnWidths[i] = 100;
        for (int i = 0; i < rows; i++)
            rowHeights[i] = 50;

        var table = slide.Shapes.AddTable(50, 50, columnWidths, rowHeights);

        if (dataArray != null)
        {
            for (int i = 0; i < Math.Min(rows, dataArray.Count); i++)
            {
                var rowArray = dataArray[i]?.AsArray();
                if (rowArray != null)
                {
                    for (int j = 0; j < Math.Min(columns, rowArray.Count); j++)
                    {
                        table[j, i].TextFrame.Text = rowArray[j]?.GetValue<string>() ?? "";
                    }
                }
            }
        }

        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"Table ({rows}x{columns}) added to slide {slideIndex}: {path}");
    }
}

