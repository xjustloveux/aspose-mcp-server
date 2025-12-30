using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Pdf;

/// <summary>
///     Tool for managing annotations in PDF documents (add, delete, edit, get)
/// </summary>
public class PdfAnnotationTool : IAsposeTool
{
    public string Description => @"Manage annotations in PDF documents. Supports 4 operations: add, delete, edit, get.

Usage examples:
- Add annotation: pdf_annotation(operation='add', path='doc.pdf', pageIndex=1, text='Note', x=100, y=100)
- Delete annotation: pdf_annotation(operation='delete', path='doc.pdf', pageIndex=1, annotationIndex=1)
- Edit annotation: pdf_annotation(operation='edit', path='doc.pdf', pageIndex=1, annotationIndex=1, text='Updated Note')
- Get annotations: pdf_annotation(operation='get', path='doc.pdf', pageIndex=1)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add an annotation (required params: path, pageIndex, text, x, y)
- 'delete': Delete annotation(s) (required params: path, pageIndex; optional: annotationIndex, deletes all if omitted)
- 'edit': Edit an annotation (required params: path, pageIndex, annotationIndex, text)
- 'get': Get all annotations (required params: path, pageIndex)",
                @enum = new[] { "add", "delete", "edit", "get" }
            },
            path = new
            {
                type = "string",
                description = "PDF file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to overwrite input)"
            },
            pageIndex = new
            {
                type = "number",
                description = "Page index (1-based, required for add, delete, edit)"
            },
            annotationIndex = new
            {
                type = "number",
                description =
                    "Annotation index (1-based, required for edit, optional for delete - deletes all if omitted)"
            },
            text = new
            {
                type = "string",
                description = "Annotation text (required for add, edit)"
            },
            x = new
            {
                type = "number",
                description =
                    "X position in points (origin is bottom-left, 72 points = 1 inch, for add/edit, default: 100)"
            },
            y = new
            {
                type = "number",
                description =
                    "Y position in points (origin is bottom-left, 72 points = 1 inch, for add/edit, default: 700)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);

        string? outputPath = null;
        if (operation.ToLower() != "get")
            outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        return operation.ToLower() switch
        {
            "add" => await AddAnnotation(path, outputPath!, arguments),
            "delete" => await DeleteAnnotation(path, outputPath!, arguments),
            "edit" => await EditAnnotation(path, outputPath!, arguments),
            "get" => await GetAnnotations(path, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds an annotation to a PDF page
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing pageIndex, text, x, y</param>
    /// <returns>Success message</returns>
    /// <exception cref="ArgumentException">Thrown when pageIndex is out of range</exception>
    private Task<string> AddAnnotation(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");
            var text = ArgumentHelper.GetString(arguments, "text");
            var x = ArgumentHelper.GetDouble(arguments, "x", "x", false, 100);
            var y = ArgumentHelper.GetDouble(arguments, "y", "y", false, 700);

            using var document = new Document(path);
            if (pageIndex < 1 || pageIndex > document.Pages.Count)
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

            var page = document.Pages[pageIndex];
            var annotation = new TextAnnotation(page, new Rectangle(x, y, x + 200, y + 50))
            {
                Title = "Comment",
                Subject = "Annotation",
                Contents = text,
                Open = false,
                Icon = TextIcon.Note
            };

            page.Annotations.Add(annotation);
            document.Save(outputPath);
            return $"Added annotation to page {pageIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes annotation(s) from a PDF page
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing pageIndex, optional annotationIndex (deletes all if omitted)</param>
    /// <returns>Success message</returns>
    /// <exception cref="ArgumentException">Thrown when pageIndex or annotationIndex is out of range</exception>
    private Task<string> DeleteAnnotation(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");
            var annotationIndex = ArgumentHelper.GetIntNullable(arguments, "annotationIndex");

            using var document = new Document(path);
            if (pageIndex < 1 || pageIndex > document.Pages.Count)
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

            var page = document.Pages[pageIndex];

            if (annotationIndex.HasValue)
            {
                if (annotationIndex.Value < 1 || annotationIndex.Value > page.Annotations.Count)
                    throw new ArgumentException($"annotationIndex must be between 1 and {page.Annotations.Count}");

                page.Annotations.Delete(annotationIndex.Value);
                document.Save(outputPath);
                return $"Deleted annotation {annotationIndex.Value} from page {pageIndex}. Output: {outputPath}";
            }

            var count = page.Annotations.Count;
            if (count == 0)
                throw new ArgumentException($"No annotations found on page {pageIndex}");

            page.Annotations.Delete();
            document.Save(outputPath);
            return $"Deleted all {count} annotation(s) from page {pageIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits an annotation in a PDF
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing pageIndex, annotationIndex, text, optional x, y</param>
    /// <returns>Success message</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when pageIndex or annotationIndex is out of range, or annotation is not a
    ///     TextAnnotation
    /// </exception>
    private Task<string> EditAnnotation(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");
            var annotationIndex = ArgumentHelper.GetInt(arguments, "annotationIndex");
            var text = ArgumentHelper.GetString(arguments, "text");
            var x = ArgumentHelper.GetDoubleNullable(arguments, "x");
            var y = ArgumentHelper.GetDoubleNullable(arguments, "y");

            using var document = new Document(path);
            if (pageIndex < 1 || pageIndex > document.Pages.Count)
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

            var page = document.Pages[pageIndex];
            if (annotationIndex < 1 || annotationIndex > page.Annotations.Count)
                throw new ArgumentException($"annotationIndex must be between 1 and {page.Annotations.Count}");

            try
            {
                var annotation = page.Annotations[annotationIndex];
                if (annotation is TextAnnotation textAnnotation)
                {
                    textAnnotation.Contents = text;
                    if (x.HasValue && y.HasValue)
                        textAnnotation.Rect = new Rectangle(x.Value, y.Value, x.Value + 200, y.Value + 50);
                }
                else
                {
                    throw new ArgumentException(
                        $"Annotation at index {annotationIndex} is not a TextAnnotation and cannot be edited");
                }
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"Failed to edit annotation: {ex.Message}");
            }

            document.Save(outputPath);
            return $"Edited annotation {annotationIndex} on page {pageIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all annotations from a PDF page
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="arguments">JSON arguments containing optional pageIndex</param>
    /// <returns>JSON string with all annotations</returns>
    /// <exception cref="ArgumentException">Thrown when pageIndex is out of range or failed to get annotations</exception>
    private Task<string> GetAnnotations(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var pageIndex = ArgumentHelper.GetIntNullable(arguments, "pageIndex");

            using var document = new Document(path);
            var annotationList = new List<object>();

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
        });
    }

    /// <summary>
    ///     Collects annotation information from a page
    /// </summary>
    /// <param name="page">The PDF page</param>
    /// <param name="pageIndex">The page index (1-based)</param>
    /// <param name="annotationList">The list to add annotation info to</param>
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