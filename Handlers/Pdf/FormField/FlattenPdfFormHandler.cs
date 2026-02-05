using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.FormField;

/// <summary>
///     Handler for flattening form fields in a PDF document,
///     converting interactive fields to static content.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class FlattenPdfFormHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "flatten";

    /// <summary>
    ///     Flattens all form fields in the PDF document.
    ///     After flattening, form fields become part of the page content and are no longer interactive.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">No additional parameters required.</param>
    /// <returns>Success message with count of flattened fields.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var document = context.Document;
        var fieldCount = document.Form.Count;

        document.Flatten();

        MarkModified(context);

        return new SuccessResult
        {
            Message = $"Flattened {fieldCount} form field(s). Fields are now static content."
        };
    }
}
