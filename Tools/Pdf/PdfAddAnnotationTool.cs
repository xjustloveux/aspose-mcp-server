using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;

namespace AsposeMcpServer.Tools;

public class PdfAddAnnotationTool : IAsposeTool
{
    public string Description => "Add an annotation (comment) to a PDF document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "PDF file path"
            },
            pageIndex = new
            {
                type = "number",
                description = "Page index (1-based)"
            },
            text = new
            {
                type = "string",
                description = "Annotation text"
            },
            x = new
            {
                type = "number",
                description = "X position (optional, default: 100)"
            },
            y = new
            {
                type = "number",
                description = "Y position (optional, default: 700)"
            }
        },
        required = new[] { "path", "pageIndex", "text" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");
        var text = arguments?["text"]?.GetValue<string>() ?? throw new ArgumentException("text is required");
        var x = arguments?["x"]?.GetValue<double>() ?? 100;
        var y = arguments?["y"]?.GetValue<double>() ?? 700;

        using var document = new Document(path);
        var page = document.Pages[pageIndex];

        var textAnnotation = new FreeTextAnnotation(page, new Rectangle(x, y, x + 200, y + 50), new DefaultAppearance())
        {
            Title = "Comment",
            Contents = text,
            Color = Aspose.Pdf.Color.Yellow
        };

        page.Annotations.Add(textAnnotation);
        document.Save(path);

        return await Task.FromResult($"Annotation added to page {pageIndex}: {path}");
    }
}

