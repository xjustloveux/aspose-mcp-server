using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Pdf.Annotation;

namespace AsposeMcpServer.Handlers.Pdf.Annotation;

/// <summary>
///     Handler for getting annotations from PDF documents.
/// </summary>
[ResultType(typeof(GetAnnotationsResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var getParams = ExtractGetParameters(parameters);

        var document = context.Document;

        var annotations = new List<AnnotationInfo>();

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

        var result = new GetAnnotationsResult
        {
            Count = annotations.Count,
            Annotations = annotations
        };

        return result;
    }

    /// <summary>
    ///     Adds annotations from a page to the collection.
    /// </summary>
    /// <param name="page">The page to get annotations from.</param>
    /// <param name="pageIndex">The page index.</param>
    /// <param name="annotations">The annotation collection to add to.</param>
    private static void AddPageAnnotations(Aspose.Pdf.Page page, int pageIndex, List<AnnotationInfo> annotations)
    {
        for (var i = 1; i <= page.Annotations.Count; i++)
        {
            var annotation = page.Annotations[i];
            var title = (annotation as MarkupAnnotation)?.Title;
            var subject = (annotation as MarkupAnnotation)?.Subject;

            annotations.Add(new AnnotationInfo
            {
                PageIndex = pageIndex,
                Index = i,
                Type = annotation.AnnotationType.ToString(),
                Title = title,
                Subject = subject,
                Contents = annotation.Contents,
                Rect = new AnnotationRect
                {
                    X = annotation.Rect.LLX,
                    Y = annotation.Rect.LLY,
                    Width = annotation.Rect.Width,
                    Height = annotation.Rect.Height
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
    private sealed record GetParameters(int PageIndex);
}
