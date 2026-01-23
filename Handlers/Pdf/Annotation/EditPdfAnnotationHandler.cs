using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.Annotation;

/// <summary>
///     Handler for editing annotations in PDF documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
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
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var editParams = ExtractEditParameters(parameters);

        var document = context.Document;

        if (editParams.PageIndex < 1 || editParams.PageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[editParams.PageIndex];

        if (editParams.AnnotationIndex < 1 || editParams.AnnotationIndex > page.Annotations.Count)
            throw new ArgumentException(
                $"annotationIndex must be between 1 and {page.Annotations.Count}");

        var annotation = page.Annotations[editParams.AnnotationIndex];

        if (!string.IsNullOrEmpty(editParams.Text))
            annotation.Contents = editParams.Text;

        if (annotation is MarkupAnnotation markupAnnotation)
        {
            if (!string.IsNullOrEmpty(editParams.Title))
                markupAnnotation.Title = editParams.Title;

            if (!string.IsNullOrEmpty(editParams.Subject))
                markupAnnotation.Subject = editParams.Subject;
        }

        MarkModified(context);

        return new SuccessResult
            { Message = $"Annotation {editParams.AnnotationIndex} on page {editParams.PageIndex} updated." };
    }

    /// <summary>
    ///     Extracts edit parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted edit parameters.</returns>
    private static EditParameters ExtractEditParameters(OperationParameters parameters)
    {
        return new EditParameters(
            parameters.GetRequired<int>("pageIndex"),
            parameters.GetRequired<int>("annotationIndex"),
            parameters.GetOptional<string?>("text"),
            parameters.GetOptional<string?>("title"),
            parameters.GetOptional<string?>("subject")
        );
    }

    /// <summary>
    ///     Record to hold edit annotation parameters.
    /// </summary>
    private sealed record EditParameters(
        int PageIndex,
        int AnnotationIndex,
        string? Text,
        string? Title,
        string? Subject);
}
