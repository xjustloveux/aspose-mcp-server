using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Annotation;

/// <summary>
///     Handler for editing annotations in PDF documents.
/// </summary>
public class EditPdfAnnotationHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "edit";

    /// <summary>
    ///     Edits an existing annotation on a page.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: pageIndex, annotationIndex.
    ///     Optional: text, title, subject.
    /// </param>
    /// <returns>Success message with edit details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var pageIndex = parameters.GetRequired<int>("pageIndex");
        var annotationIndex = parameters.GetRequired<int>("annotationIndex");
        var text = parameters.GetOptional<string?>("text");
        var title = parameters.GetOptional<string?>("title");
        var subject = parameters.GetOptional<string?>("subject");

        var document = context.Document;

        if (pageIndex < 1 || pageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[pageIndex];

        if (annotationIndex < 1 || annotationIndex > page.Annotations.Count)
            throw new ArgumentException(
                $"annotationIndex must be between 1 and {page.Annotations.Count}");

        var annotation = page.Annotations[annotationIndex];

        if (!string.IsNullOrEmpty(text))
            annotation.Contents = text;

        if (annotation is MarkupAnnotation markupAnnotation)
        {
            if (!string.IsNullOrEmpty(title))
                markupAnnotation.Title = title;

            if (!string.IsNullOrEmpty(subject))
                markupAnnotation.Subject = subject;
        }

        MarkModified(context);

        return Success($"Annotation {annotationIndex} on page {pageIndex} updated.");
    }
}
