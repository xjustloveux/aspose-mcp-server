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
- 'delete': Delete an annotation (required params: path, pageIndex, annotationIndex)
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
                description = "Annotation index (1-based, required for delete, edit)"
            },
            text = new
            {
                type = "string",
                description = "Annotation text (required for add, edit)"
            },
            x = new
            {
                type = "number",
                description = "X position (for add, edit, default: 100)"
            },
            y = new
            {
                type = "number",
                description = "Y position (for add, edit, default: 700)"
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

        // Only get outputPath for operations that modify the document
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
    ///     Deletes an annotation from a PDF page
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing pageIndex, annotationIndex</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteAnnotation(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");
            var annotationIndex = ArgumentHelper.GetInt(arguments, "annotationIndex");

            using var document = new Document(path);
            if (pageIndex < 1 || pageIndex > document.Pages.Count)
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

            var page = document.Pages[pageIndex];
            if (annotationIndex < 1 || annotationIndex > page.Annotations.Count)
                throw new ArgumentException($"annotationIndex must be between 1 and {page.Annotations.Count}");

            page.Annotations.Delete(annotationIndex);
            document.Save(outputPath);
            return $"Deleted annotation {annotationIndex} from page {pageIndex}. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits an annotation in a PDF
    /// </summary>
    /// <param name="path">Input file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing pageIndex, annotationIndex, text, optional x, y</param>
    /// <returns>Success message</returns>
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
    private Task<string> GetAnnotations(string path, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var pageIndex = ArgumentHelper.GetIntNullable(arguments, "pageIndex");

            try
            {
                using var document = new Document(path);
                var annotationList = new List<object>();

                if (pageIndex.HasValue)
                {
                    if (pageIndex.Value < 1 || pageIndex.Value > document.Pages.Count)
                        throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

                    var page = document.Pages[pageIndex.Value];
                    for (var i = 1; i <= page.Annotations.Count; i++)
                        try
                        {
                            var annotation = page.Annotations[i];
                            var annotationInfo = new Dictionary<string, object?>
                            {
                                ["index"] = i,
                                ["pageIndex"] = pageIndex.Value,
                                ["type"] = annotation.GetType().Name
                            };
                            if (annotation is TextAnnotation textAnnotation)
                                annotationInfo["text"] = textAnnotation.Contents ?? "";
                            if (annotation.Rect != null)
                            {
                                annotationInfo["x"] = annotation.Rect.LLX;
                                annotationInfo["y"] = annotation.Rect.LLY;
                            }

                            annotationList.Add(annotationInfo);
                        }
                        catch (Exception ex)
                        {
                            annotationList.Add(new { index = i, pageIndex = pageIndex.Value, error = ex.Message });
                        }

                    var result = new
                    {
                        count = annotationList.Count,
                        pageIndex = pageIndex.Value,
                        items = annotationList
                    };
                    return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
                }
                else
                {
                    for (var i = 1; i <= document.Pages.Count; i++)
                        try
                        {
                            var page = document.Pages[i];
                            for (var j = 1; j <= page.Annotations.Count; j++)
                                try
                                {
                                    var annotation = page.Annotations[j];
                                    var annotationInfo = new Dictionary<string, object?>
                                    {
                                        ["index"] = j,
                                        ["pageIndex"] = i,
                                        ["type"] = annotation.GetType().Name
                                    };
                                    if (annotation is TextAnnotation textAnnotation)
                                        annotationInfo["text"] = textAnnotation.Contents ?? "";
                                    if (annotation.Rect != null)
                                    {
                                        annotationInfo["x"] = annotation.Rect.LLX;
                                        annotationInfo["y"] = annotation.Rect.LLY;
                                    }

                                    annotationList.Add(annotationInfo);
                                }
                                catch (Exception ex)
                                {
                                    annotationList.Add(new { index = j, pageIndex = i, error = ex.Message });
                                }
                        }
                        catch (Exception ex)
                        {
                            annotationList.Add(new { pageIndex = i, error = ex.Message });
                        }

                    if (annotationList.Count == 0)
                    {
                        var emptyResult = new
                        {
                            count = 0,
                            items = Array.Empty<object>(),
                            message = "No annotations found"
                        };
                        return JsonSerializer.Serialize(emptyResult,
                            new JsonSerializerOptions { WriteIndented = true });
                    }

                    var result = new
                    {
                        count = annotationList.Count,
                        items = annotationList
                    };
                    return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
                }
            }
            catch (Exception ex)
            {
                throw new ArgumentException($"Failed to get annotations: {ex.Message}");
            }
        });
    }
}