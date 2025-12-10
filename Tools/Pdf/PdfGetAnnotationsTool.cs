using System.Text.Json.Nodes;
using System.Text;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;

namespace AsposeMcpServer.Tools;

public class PdfGetAnnotationsTool : IAsposeTool
{
    public string Description => "Get all annotations from PDF document or specific page";

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
                description = "Page index (1-based, optional, if not provided returns all annotations)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var pageIndex = arguments?["pageIndex"]?.GetValue<int?>();

        using var document = new Document(path);
        var sb = new StringBuilder();

        if (pageIndex.HasValue)
        {
            if (pageIndex.Value < 1 || pageIndex.Value > document.Pages.Count)
            {
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");
            }

            var page = document.Pages[pageIndex.Value];
            sb.AppendLine($"Annotations on Page {pageIndex.Value} ({page.Annotations.Count}):");
            sb.AppendLine();

            for (int i = 0; i < page.Annotations.Count; i++)
            {
                var annotation = page.Annotations[i];
                sb.AppendLine($"[{i}] {annotation.GetType().Name}");
                if (annotation is MarkupAnnotation markup)
                {
                    sb.AppendLine($"    Title: {markup.Title ?? "(none)"}");
                    sb.AppendLine($"    Contents: {markup.Contents ?? "(none)"}");
                }
                sb.AppendLine($"    Color: {annotation.Color}");
                sb.AppendLine();
            }
        }
        else
        {
            var totalCount = 0;
            for (int p = 1; p <= document.Pages.Count; p++)
            {
                var page = document.Pages[p];
                if (page.Annotations.Count > 0)
                {
                    sb.AppendLine($"Page {p} ({page.Annotations.Count} annotations):");
                    for (int i = 0; i < page.Annotations.Count; i++)
                    {
                        var annotation = page.Annotations[i];
                        sb.AppendLine($"  [{i}] {annotation.GetType().Name}");
                        if (annotation is MarkupAnnotation markup)
                        {
                            sb.AppendLine($"      Title: {markup.Title ?? "(none)"}");
                            sb.AppendLine($"      Contents: {markup.Contents ?? "(none)"}");
                        }
                    }
                    sb.AppendLine();
                    totalCount += page.Annotations.Count;
                }
            }
            sb.Insert(0, $"Total Annotations: {totalCount}\n\n");
        }

        return await Task.FromResult(sb.ToString());
    }
}

