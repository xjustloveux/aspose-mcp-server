using System.Text;
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
- Delete annotation: pdf_annotation(operation='delete', path='doc.pdf', pageIndex=1, annotationIndex=0)
- Edit annotation: pdf_annotation(operation='edit', path='doc.pdf', pageIndex=1, annotationIndex=0, text='Updated Note')
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
                description = "Annotation index (0-based, required for delete, edit)"
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

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");

        return operation.ToLower() switch
        {
            "add" => await AddAnnotation(arguments),
            "delete" => await DeleteAnnotation(arguments),
            "edit" => await EditAnnotation(arguments),
            "get" => await GetAnnotations(arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds an annotation to a PDF page
    /// </summary>
    /// <param name="arguments">
    ///     JSON arguments containing path, pageIndex, annotationType, x, y, width, height, optional text,
    ///     outputPath
    /// </param>
    /// <returns>Success message</returns>
    private async Task<string> AddAnnotation(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
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
        return await Task.FromResult($"Successfully added annotation to page {pageIndex}. Output: {outputPath}");
    }

    /// <summary>
    ///     Deletes an annotation from a PDF page
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, pageIndex, annotationIndex, optional outputPath</param>
    /// <returns>Success message</returns>
    private async Task<string> DeleteAnnotation(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");
        var annotationIndex = ArgumentHelper.GetInt(arguments, "annotationIndex");

        SecurityHelper.ValidateFilePath(path);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        if (annotationIndex < 0 || annotationIndex >= page.Annotations.Count)
            throw new ArgumentException($"annotationIndex must be between 0 and {page.Annotations.Count - 1}");

        page.Annotations.Delete(annotationIndex);
        document.Save(outputPath);
        return await Task.FromResult(
            $"Successfully deleted annotation {annotationIndex} from page {pageIndex}. Output: {outputPath}");
    }

    /// <summary>
    ///     Edits an annotation in a PDF
    /// </summary>
    /// <param name="arguments">
    ///     JSON arguments containing path, pageIndex, annotationIndex, optional text, x, y, width, height,
    ///     outputPath
    /// </param>
    /// <returns>Success message</returns>
    private async Task<string> EditAnnotation(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var pageIndex = ArgumentHelper.GetInt(arguments, "pageIndex");
        var annotationIndex = ArgumentHelper.GetInt(arguments, "annotationIndex");
        var text = ArgumentHelper.GetString(arguments, "text");
        var x = ArgumentHelper.GetDoubleNullable(arguments, "x");
        var y = ArgumentHelper.GetDoubleNullable(arguments, "y");

        SecurityHelper.ValidateFilePath(path);
        SecurityHelper.ValidateFilePath(outputPath, "outputPath");

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];
        if (annotationIndex < 0 || annotationIndex >= page.Annotations.Count)
            throw new ArgumentException($"annotationIndex must be between 0 and {page.Annotations.Count - 1}");

        var annotation = page.Annotations[annotationIndex];
        if (annotation is TextAnnotation textAnnotation)
        {
            textAnnotation.Contents = text;
            if (x.HasValue && y.HasValue)
                textAnnotation.Rect = new Rectangle(x.Value, y.Value, x.Value + 200, y.Value + 50);
        }

        document.Save(outputPath);
        return await Task.FromResult(
            $"Successfully edited annotation {annotationIndex} on page {pageIndex}. Output: {outputPath}");
    }

    /// <summary>
    ///     Gets all annotations from a PDF page
    /// </summary>
    /// <param name="arguments">JSON arguments containing path, pageIndex</param>
    /// <returns>Formatted string with all annotations</returns>
    private async Task<string> GetAnnotations(JsonObject? arguments)
    {
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var pageIndex = ArgumentHelper.GetIntNullable(arguments, "pageIndex");

        SecurityHelper.ValidateFilePath(path);

        using var document = new Document(path);
        var sb = new StringBuilder();
        sb.AppendLine("=== PDF Annotations ===");
        sb.AppendLine();

        if (pageIndex.HasValue)
        {
            if (pageIndex.Value < 1 || pageIndex.Value > document.Pages.Count)
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

            var page = document.Pages[pageIndex.Value];
            sb.AppendLine($"Page {pageIndex.Value} Annotations ({page.Annotations.Count}):");
            for (var i = 0; i < page.Annotations.Count; i++)
            {
                var annotation = page.Annotations[i];
                sb.AppendLine($"  [{i}] Type: {annotation.GetType().Name}");
                if (annotation is TextAnnotation textAnnotation)
                    sb.AppendLine($"      Text: {textAnnotation.Contents}");
                sb.AppendLine($"      Position: ({annotation.Rect.LLX}, {annotation.Rect.LLY})");
                sb.AppendLine();
            }
        }
        else
        {
            var totalCount = 0;
            for (var i = 1; i <= document.Pages.Count; i++)
            {
                var page = document.Pages[i];
                if (page.Annotations.Count > 0)
                {
                    sb.AppendLine($"Page {i} ({page.Annotations.Count} annotations):");
                    for (var j = 0; j < page.Annotations.Count; j++)
                    {
                        var annotation = page.Annotations[j];
                        sb.AppendLine($"  [{j}] Type: {annotation.GetType().Name}");
                        if (annotation is TextAnnotation textAnnotation)
                            sb.AppendLine($"      Text: {textAnnotation.Contents}");
                    }

                    sb.AppendLine();
                    totalCount += page.Annotations.Count;
                }
            }

            sb.AppendLine($"Total Annotations: {totalCount}");
        }

        return await Task.FromResult(sb.ToString());
    }
}