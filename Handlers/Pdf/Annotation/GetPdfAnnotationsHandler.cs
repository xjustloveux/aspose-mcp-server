using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Annotation;

/// <summary>
///     Handler for getting annotations from PDF documents.
/// </summary>
public class GetPdfAnnotationsHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Gets all annotations from a specific page.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: pageIndex (default: 1, or 0 for all pages).
    /// </param>
    /// <returns>JSON result with annotation information.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var getParams = ExtractGetParameters(parameters);

        var document = context.Document;

        var annotations = new List<object>();

        if (getParams.PageIndex == 0)
        {
            for (var p = 1; p <= document.Pages.Count; p++) AddPageAnnotations(document.Pages[p], p, annotations);
        }
        else
        {
            if (getParams.PageIndex < 1 || getParams.PageIndex > document.Pages.Count)
                throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

            AddPageAnnotations(document.Pages[getParams.PageIndex], getParams.PageIndex, annotations);
        }

        var result = new
        {
            count = annotations.Count,
            annotations
        };

        return JsonResult(result);
    }

    /// <summary>
    ///     Adds annotations from a page to the collection.
    /// </summary>
    /// <param name="page">The page to get annotations from.</param>
    /// <param name="pageIndex">The page index.</param>
    /// <param name="annotations">The annotation collection to add to.</param>
    private static void AddPageAnnotations(Aspose.Pdf.Page page, int pageIndex, List<object> annotations)
    {
        for (var i = 1; i <= page.Annotations.Count; i++)
        {
            var annotation = page.Annotations[i];
            var title = (annotation as MarkupAnnotation)?.Title;
            var subject = (annotation as MarkupAnnotation)?.Subject;

            annotations.Add(new
            {
                pageIndex,
                index = i,
                type = annotation.AnnotationType.ToString(),
                title,
                subject,
                contents = annotation.Contents,
                rect = new
                {
                    x = annotation.Rect.LLX,
                    y = annotation.Rect.LLY,
                    width = annotation.Rect.Width,
                    height = annotation.Rect.Height
                }
            });
        }
    }

    /// <summary>
    ///     Extracts get parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted get parameters.</returns>
    private static GetParameters ExtractGetParameters(OperationParameters parameters)
    {
        return new GetParameters(
            parameters.GetOptional("pageIndex", 1)
        );
    }

    /// <summary>
    ///     Record to hold get annotations parameters.
    /// </summary>
    private record GetParameters(int PageIndex);
}
