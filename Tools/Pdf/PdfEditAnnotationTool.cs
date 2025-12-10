using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;

namespace AsposeMcpServer.Tools;

public class PdfEditAnnotationTool : IAsposeTool
{
    public string Description => "Edit annotation properties in PDF document";

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
            annotationIndex = new
            {
                type = "number",
                description = "Annotation index on the page (0-based)"
            },
            text = new
            {
                type = "string",
                description = "New annotation text (optional)"
            },
            title = new
            {
                type = "string",
                description = "New annotation title (optional)"
            },
            color = new
            {
                type = "string",
                description = "Color name (e.g., 'Yellow', 'Red', optional)"
            }
        },
        required = new[] { "path", "pageIndex", "annotationIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");
        var annotationIndex = arguments?["annotationIndex"]?.GetValue<int>() ?? throw new ArgumentException("annotationIndex is required");
        var text = arguments?["text"]?.GetValue<string>();
        var title = arguments?["title"]?.GetValue<string>();
        var color = arguments?["color"]?.GetValue<string>();

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
        {
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");
        }

        var page = document.Pages[pageIndex];
        if (annotationIndex < 0 || annotationIndex >= page.Annotations.Count)
        {
            throw new ArgumentException($"annotationIndex must be between 0 and {page.Annotations.Count - 1}");
        }

        var annotation = page.Annotations[annotationIndex];

        if (!string.IsNullOrEmpty(text) && annotation is MarkupAnnotation markupAnnotation)
        {
            markupAnnotation.Contents = text;
        }

        if (!string.IsNullOrEmpty(title) && annotation is MarkupAnnotation titleAnnotation)
        {
            titleAnnotation.Title = title;
        }

        if (!string.IsNullOrEmpty(color))
        {
            var colorObj = System.Drawing.Color.FromName(color);
            if (colorObj.IsKnownColor)
            {
                annotation.Color = Aspose.Pdf.Color.FromRgb(colorObj.R, colorObj.G, colorObj.B);
            }
        }

        document.Save(path);
        return await Task.FromResult($"Annotation {annotationIndex} updated on page {pageIndex}: {path}");
    }
}

