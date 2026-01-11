using Aspose.Pdf;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Signature;

/// <summary>
///     Handler for deleting signatures from PDF documents.
/// </summary>
public class DeletePdfSignatureHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "delete";

    /// <summary>
    ///     Deletes a signature field from the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: signatureName or signatureIndex.
    /// </param>
    /// <returns>Success message with deletion details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var signatureName = parameters.GetOptional<string?>("signatureName");
        var signatureIndex = parameters.GetOptional<int?>("signatureIndex");

        if (string.IsNullOrEmpty(signatureName) && !signatureIndex.HasValue)
            throw new ArgumentException("Either 'signatureName' or 'signatureIndex' is required");

        var document = context.Document;

        var signatureFields = document.Form.Fields
            .OfType<SignatureField>()
            .ToList();

        if (signatureFields.Count == 0)
            throw new ArgumentException("Document has no signature fields");

        SignatureField? fieldToDelete = null;

        if (!string.IsNullOrEmpty(signatureName))
        {
            fieldToDelete = signatureFields.FirstOrDefault(f => f.PartialName == signatureName);
            if (fieldToDelete == null)
                throw new ArgumentException($"Signature field '{signatureName}' not found");
        }
        else if (signatureIndex.HasValue)
        {
            if (signatureIndex.Value < 0 || signatureIndex.Value >= signatureFields.Count)
                throw new ArgumentException(
                    $"signatureIndex must be between 0 and {signatureFields.Count - 1}");
            fieldToDelete = signatureFields[signatureIndex.Value];
        }

        if (fieldToDelete != null)
        {
            document.Form.Delete(fieldToDelete.PartialName);
            MarkModified(context);
            return Success($"Signature '{fieldToDelete.PartialName}' deleted.");
        }

        throw new ArgumentException("Could not find signature to delete");
    }
}
