using System.Text.Json.Nodes;
using Aspose.Pdf;

namespace AsposeMcpServer.Tools;

public class PdfDeleteAnnotationTool : IAsposeTool
{
    public string Description => "Delete annotation(s) from PDF document";

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
                description = "Annotation index to delete (0-based)"
            },
            annotationIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Array of annotation indices to delete (0-based, optional, overrides annotationIndex)"
            }
        },
        required = new[] { "path", "pageIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");
        var annotationIndex = arguments?["annotationIndex"]?.GetValue<int?>();
        var annotationIndicesArray = arguments?["annotationIndices"]?.AsArray();

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
        {
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");
        }

        var page = document.Pages[pageIndex];
        var annotations = page.Annotations;

        List<int> annotationsToDelete;
        if (annotationIndicesArray != null && annotationIndicesArray.Count > 0)
        {
            annotationsToDelete = annotationIndicesArray.Select(a => a?.GetValue<int>()).Where(a => a.HasValue).Select(a => a!.Value).OrderByDescending(a => a).ToList();
        }
        else if (annotationIndex.HasValue)
        {
            annotationsToDelete = new List<int> { annotationIndex.Value };
        }
        else
        {
            throw new ArgumentException("Either annotationIndex or annotationIndices must be provided");
        }

        foreach (var index in annotationsToDelete)
        {
            if (index < 0 || index >= annotations.Count)
            {
                continue;
            }
            annotations.Delete(index);
        }

        document.Save(path);
        return await Task.FromResult($"Deleted {annotationsToDelete.Count} annotation(s) from page {pageIndex}. Remaining annotations: {annotations.Count}");
    }
}

