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
        var p = ExtractDeleteParameters(parameters);

        if (string.IsNullOrEmpty(p.SignatureName) && !p.SignatureIndex.HasValue)
            throw new ArgumentException("Either 'signatureName' or 'signatureIndex' is required");

        var document = context.Document;

        var signatureFields = document.Form.Fields
            .OfType<SignatureField>()
            .ToList();

        if (signatureFields.Count == 0)
            throw new ArgumentException("Document has no signature fields");

        SignatureField? fieldToDelete = null;

        if (!string.IsNullOrEmpty(p.SignatureName))
        {
            fieldToDelete = signatureFields.FirstOrDefault(f => f.PartialName == p.SignatureName);
            if (fieldToDelete == null)
                throw new ArgumentException($"Signature field '{p.SignatureName}' not found");
        }
        else if (p.SignatureIndex.HasValue)
        {
            if (p.SignatureIndex.Value < 0 || p.SignatureIndex.Value >= signatureFields.Count)
                throw new ArgumentException(
                    $"signatureIndex must be between 0 and {signatureFields.Count - 1}");
            fieldToDelete = signatureFields[p.SignatureIndex.Value];
        }

        if (fieldToDelete != null)
        {
            document.Form.Delete(fieldToDelete.PartialName);
            MarkModified(context);
            return Success($"Signature '{fieldToDelete.PartialName}' deleted.");
        }

        throw new ArgumentException("Could not find signature to delete");
    }

    /// <summary>
    ///     Extracts parameters for delete operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static DeleteParameters ExtractDeleteParameters(OperationParameters parameters)
    {
        return new DeleteParameters(
            parameters.GetOptional<string?>("signatureName"),
            parameters.GetOptional<int?>("signatureIndex")
        );
    }

    /// <summary>
    ///     Parameters for delete operation.
    /// </summary>
    /// <param name="SignatureName">The optional signature name.</param>
    /// <param name="SignatureIndex">The optional 0-based signature index.</param>
    private record DeleteParameters(string? SignatureName, int? SignatureIndex);
}
