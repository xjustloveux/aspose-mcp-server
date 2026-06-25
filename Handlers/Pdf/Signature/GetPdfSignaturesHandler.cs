using Aspose.Pdf;
using Aspose.Pdf.Facades;
using Aspose.Pdf.Forms;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Pdf.Signature;

namespace AsposeMcpServer.Handlers.Pdf.Signature;

/// <summary>
///     Handler for retrieving digital signatures from PDF documents.
/// </summary>
[ResultType(typeof(GetSignaturesResult))]
public class GetPdfSignaturesHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "get";

    /// <summary>
    ///     Retrieves all digital signatures from the PDF document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">No parameters required.</param>
    /// <returns>A JSON string containing signature information.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var document = context.Document;

        // PdfFileSignature is intentionally NOT wrapped in 'using': disposing it disposes the bound
        // document, which is owned by the caller/session and reused across operations. get must be
        // read-only and leave the document usable.
        var pdfSign = new PdfFileSignature(document);

        // Enumerate the signature FIELDS (signed or not) in the same order delete indexes them
        // (document.Form.Fields.OfType<SignatureField>()), so a get-reported index round-trips into
        // delete's signatureIndex. GetSignNames() lists only SIGNED signatures, which diverged from
        // delete's all-fields index space and made the get-reported index unusable for delete.
        var signatureFields = document.Form.Fields.OfType<SignatureField>().ToList();

        if (signatureFields.Count == 0)
            return new GetSignaturesResult
            {
                Count = 0,
                Items = [],
                Message = "No signatures found"
            };

        List<SignatureInfo> signatureList = [];
        for (var i = 0; i < signatureFields.Count; i++)
        {
            var field = signatureFields[i];
            bool isValid;
            bool hasCertificate;

            try
            {
                isValid = pdfSign.VerifySignature(field.FullName);
            }
            catch
            {
                isValid = false;
            }

            try
            {
                _ = pdfSign.ExtractCertificate(field.FullName);
                hasCertificate = true;
            }
            catch
            {
                hasCertificate = false;
            }

            signatureList.Add(new SignatureInfo
            {
                Index = i,
                Name = field.PartialName,
                IsValid = isValid,
                HasCertificate = hasCertificate
            });
        }

        return new GetSignaturesResult
        {
            Count = signatureList.Count,
            Items = signatureList
        };
    }
}
