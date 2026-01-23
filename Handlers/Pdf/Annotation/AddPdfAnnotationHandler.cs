using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.Annotation;

/// <summary>
///     Handler for adding text annotations to PDF documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class AddPdfAnnotationHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a new text annotation to the specified page.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: pageIndex, text
    ///     Optional: x (default: 100), y (default: 700)
    /// </param>
    /// <returns>Success message with annotation details.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var addParams = ExtractAddParameters(parameters);

        var document = context.Document;

        if (addParams.PageIndex < 1 || addParams.PageIndex > document.Pages.Count)
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");

        var page = document.Pages[addParams.PageIndex];
        var annotation =
            new TextAnnotation(page, new Rectangle(addParams.X, addParams.Y, addParams.X + 200, addParams.Y + 50))
            {
                Title = "Comment",
                Subject = "Annotation",
                Contents = addParams.Text,
                Open = false,
                Icon = TextIcon.Note
            };

        page.Annotations.Add(annotation);

        MarkModified(context);

        return new SuccessResult { Message = $"Added annotation to page {addParams.PageIndex}." };
    }

    /// <summary>
    ///     Extracts add parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted add parameters.</returns>
    private static AddParameters ExtractAddParameters(OperationParameters parameters)
    {
        return new AddParameters(
            parameters.GetRequired<int>("pageIndex"),
            parameters.GetRequired<string>("text"),
            parameters.GetOptional("x", 100.0),
            parameters.GetOptional("y", 700.0)
        );
    }

    /// <summary>
    ///     Record to hold add annotation parameters.
    /// </summary>
    private sealed record AddParameters(int PageIndex, string Text, double X, double Y);
}
