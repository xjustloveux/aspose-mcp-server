using System.ComponentModel;
using System.Text.Json;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing annotations in PDF documents (add, delete, edit, get)
/// </summary>
[McpServerToolType]
public class PdfAnnotationTool
{
    /// <summary>
    ///     The document session manager for managing in-memory document sessions.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PdfAnnotationTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    public PdfAnnotationTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "pdf_annotation")]
    [Description(@"Manage annotations in PDF documents. Supports 4 operations: add, delete, edit, get.

Usage examples:
- Add annotation: pdf_annotation(operation='add', path='doc.pdf', pageIndex=1, text='Note', x=100, y=100)
- Delete annotation: pdf_annotation(operation='delete', path='doc.pdf', pageIndex=1, annotationIndex=1)
- Edit annotation: pdf_annotation(operation='edit', path='doc.pdf', pageIndex=1, annotationIndex=1, text='Updated Note')
- Get annotations: pdf_annotation(operation='get', path='doc.pdf', pageIndex=1)")]
    public string Execute(
        [Description(@"Operation to perform.
- 'add': Add an annotation (required params: path, pageIndex, text, x, y)
- 'delete': Delete annotation(s) (required params: path, pageIndex; optional: annotationIndex, deletes all if omitted)
- 'edit': Edit an annotation (required params: path, pageIndex, annotationIndex, text)
- 'get': Get all annotations (required params: path, pageIndex)")]
        string operation,
        [Description("PDF file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Page index (1-based, required for add, delete, edit)")]
        int? pageIndex = null,
        [Description("Annotation index (1-based, required for edit, optional for delete - deletes all if omitted)")]
        int? annotationIndex = null,
        [Description("Annotation text (required for add, edit)")]
        string? text = null,
        [Description("X position in points (origin is bottom-left, 72 points = 1 inch, for add/edit, default: 100)")]
        double x = 100,
        [Description("Y position in points (origin is bottom-left, 72 points = 1 inch, for add/edit, default: 700)")]
        double y = 700)
    {
        return operation.ToLower() switch
        {
            "add" => AddAnnotation(sessionId, path, outputPath, pageIndex, text, x, y),
            "delete" => DeleteAnnotation(sessionId, path, outputPath, pageIndex, annotationIndex),
            "edit" => EditAnnotation(sessionId, path, outputPath, pageIndex, annotationIndex, text, x, y),
            "get" => GetAnnotations(sessionId, path, pageIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a new text annotation to the specified page.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="pageIndex">The 1-based page index.</param>
    /// <param name="text">The annotation text content.</param>
    /// <param name="x">The X position in PDF coordinates.</param>
    /// <param name="y">The Y position in PDF coordinates.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    private string AddAnnotation(string? sessionId, string? path, string? outputPath, int? pageIndex, string? text,
        double x, double y)
    {
        if (!pageIndex.HasValue)
            throw new ArgumentException("pageIndex is required for add operation");
        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("text is required for add operation");

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);
        var document = ctx.Document;

        if (pageIndex.Value < 1 || pageIndex.Value > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex.Value];
        var annotation = new TextAnnotation(page, new Rectangle(x, y, x + 200, y + 50))
        {
            Title = "Comment",
            Subject = "Annotation",
            Contents = text,
            Open = false,
            Icon = TextIcon.Note
        };

        page.Annotations.Add(annotation);

        ctx.Save(outputPath);
        return $"Added annotation to page {pageIndex.Value}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Deletes one or all annotations from the specified page.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="pageIndex">The 1-based page index.</param>
    /// <param name="annotationIndex">The 1-based annotation index, or null to delete all annotations.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    private string DeleteAnnotation(string? sessionId, string? path, string? outputPath, int? pageIndex,
        int? annotationIndex)
    {
        if (!pageIndex.HasValue)
            throw new ArgumentException("pageIndex is required for delete operation");

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);
        var document = ctx.Document;

        if (pageIndex.Value < 1 || pageIndex.Value > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex.Value];

        if (annotationIndex.HasValue)
        {
            if (annotationIndex.Value < 1 || annotationIndex.Value > page.Annotations.Count)
                throw new ArgumentException($"annotationIndex must be between 1 and {page.Annotations.Count}");

            page.Annotations.Delete(annotationIndex.Value);
            ctx.Save(outputPath);
            return
                $"Deleted annotation {annotationIndex.Value} from page {pageIndex.Value}. {ctx.GetOutputMessage(outputPath)}";
        }

        var count = page.Annotations.Count;
        if (count == 0)
            throw new ArgumentException($"No annotations found on page {pageIndex.Value}");

        page.Annotations.Delete();
        ctx.Save(outputPath);
        return $"Deleted all {count} annotation(s) from page {pageIndex.Value}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Edits an existing text annotation on the specified page.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="pageIndex">The 1-based page index.</param>
    /// <param name="annotationIndex">The 1-based annotation index.</param>
    /// <param name="text">The new annotation text content.</param>
    /// <param name="x">The new X position in PDF coordinates.</param>
    /// <param name="y">The new Y position in PDF coordinates.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing or invalid.</exception>
    private string EditAnnotation(string? sessionId, string? path, string? outputPath, int? pageIndex,
        int? annotationIndex, string? text, double x, double y)
    {
        if (!pageIndex.HasValue)
            throw new ArgumentException("pageIndex is required for edit operation");
        if (!annotationIndex.HasValue)
            throw new ArgumentException("annotationIndex is required for edit operation");
        if (string.IsNullOrEmpty(text))
            throw new ArgumentException("text is required for edit operation");

        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);
        var document = ctx.Document;

        if (pageIndex.Value < 1 || pageIndex.Value > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex.Value];
        if (annotationIndex.Value < 1 || annotationIndex.Value > page.Annotations.Count)
            throw new ArgumentException($"annotationIndex must be between 1 and {page.Annotations.Count}");

        try
        {
            var annotation = page.Annotations[annotationIndex.Value];
            if (annotation is TextAnnotation textAnnotation)
            {
                textAnnotation.Contents = text;
                if (Math.Abs(x - 100) > 0.001 || Math.Abs(y - 700) > 0.001)
                    textAnnotation.Rect = new Rectangle(x, y, x + 200, y + 50);
            }
            else
            {
                throw new ArgumentException(
                    $"Annotation at index {annotationIndex.Value} is not a TextAnnotation and cannot be edited");
            }
        }
        catch (Exception ex) when (ex is not ArgumentException)
        {
            throw new ArgumentException($"Failed to edit annotation: {ex.Message}");
        }

        ctx.Save(outputPath);
        return
            $"Edited annotation {annotationIndex.Value} on page {pageIndex.Value}. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Retrieves annotations from the specified page or all pages.
    /// </summary>
    /// <param name="sessionId">The session ID for in-memory editing.</param>
    /// <param name="path">The PDF file path.</param>
    /// <param name="pageIndex">The 1-based page index, or null for all pages.</param>
    /// <returns>A JSON string containing annotation information.</returns>
    /// <exception cref="ArgumentException">Thrown when the page index is out of range.</exception>
    private string GetAnnotations(string? sessionId, string? path, int? pageIndex)
    {
        using var ctx = DocumentContext<Document>.Create(_sessionManager, sessionId, path);
        var document = ctx.Document;
        List<object> annotationList = [];

        if (pageIndex.HasValue)
        {
            if (pageIndex.Value < 1 || pageIndex.Value > document.Pages.Count)
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

            var page = document.Pages[pageIndex.Value];
            CollectAnnotationsFromPage(page, pageIndex.Value, annotationList);

            var result = new
            {
                count = annotationList.Count,
                pageIndex = pageIndex.Value,
                items = annotationList
            };
            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        }

        for (var i = 1; i <= document.Pages.Count; i++)
        {
            var page = document.Pages[i];
            CollectAnnotationsFromPage(page, i, annotationList);
        }

        var allResult = new
        {
            count = annotationList.Count,
            items = annotationList,
            message = annotationList.Count == 0 ? "No annotations found" : null
        };
        return JsonSerializer.Serialize(allResult, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Collects annotation information from a single page.
    /// </summary>
    /// <param name="page">The PDF page to collect annotations from.</param>
    /// <param name="pageIndex">The 1-based page index for reference.</param>
    /// <param name="annotationList">The list to add annotation information to.</param>
    private static void CollectAnnotationsFromPage(Page page, int pageIndex, List<object> annotationList)
    {
        for (var i = 1; i <= page.Annotations.Count; i++)
        {
            var annotation = page.Annotations[i];
            var annotationInfo = new Dictionary<string, object?>
            {
                ["index"] = i,
                ["pageIndex"] = pageIndex,
                ["type"] = annotation.GetType().Name,
                ["name"] = !string.IsNullOrEmpty(annotation.Name) ? annotation.Name : null,
                ["modified"] = annotation.Modified != DateTime.MinValue ? annotation.Modified.ToString("o") : null
            };

            if (annotation is MarkupAnnotation markupAnnotation)
            {
                annotationInfo["contents"] = markupAnnotation.Contents;
                annotationInfo["subject"] =
                    !string.IsNullOrEmpty(markupAnnotation.Subject) ? markupAnnotation.Subject : null;
                annotationInfo["title"] = !string.IsNullOrEmpty(markupAnnotation.Title) ? markupAnnotation.Title : null;
            }

            if (annotation.Rect != null)
            {
                annotationInfo["x"] = annotation.Rect.LLX;
                annotationInfo["y"] = annotation.Rect.LLY;
                annotationInfo["width"] = annotation.Rect.Width;
                annotationInfo["height"] = annotation.Rect.Height;
            }

            annotationList.Add(annotationInfo);
        }
    }
}