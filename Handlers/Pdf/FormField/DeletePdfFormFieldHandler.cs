using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.FormField;

/// <summary>
///     Handler for deleting form fields from PDF documents.
/// </summary>
public class DeletePdfFormFieldHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a form field from the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: fieldName
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var fieldName = parameters.GetRequired<string>("fieldName");

        var document = context.Document;

        if (document.Form.Cast<Field>().All(f => f.PartialName != fieldName))
            throw new ArgumentException($"Form field '{fieldName}' not found");

        document.Form.Delete(fieldName);
        MarkModified(context);

        return Success($"Deleted form field '{fieldName}'.");
    }
}
