using Aspose.Pdf;
using Aspose.Pdf.Facades;
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
        using var pdfSign = new PdfFileSignature(document);
        var signatureNames = pdfSign.GetSignNames();

        if (signatureNames.Count == 0)
            return new GetSignaturesResult
            {
                Count = 0,
                Items = [],
                Message = "No signatures found"
            };

        List<SignatureInfo> signatureList = [];
        for (var i = 0; i < signatureNames.Count; i++)
        {
            var signatureName = signatureNames[i];
            bool isValid;
            bool hasCertificate;

            try
            {
                isValid = pdfSign.VerifySignature(signatureName);
            }
            catch
            {
                isValid = false;
            }

            try
            {
                _ = pdfSign.ExtractCertificate(signatureName);
                hasCertificate = true;
            }
            catch
            {
                hasCertificate = false;
            }

            signatureList.Add(new SignatureInfo
            {
                Index = i,
                Name = signatureName,
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
