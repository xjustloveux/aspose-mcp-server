using Aspose.Pdf;
using Aspose.Pdf.Facades;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.Signature;

/// <summary>
///     Handler for retrieving digital signatures from PDF documents.
/// </summary>
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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var document = context.Document;
        using var pdfSign = new PdfFileSignature(document);
        var signatureNames = pdfSign.GetSignNames();

        if (signatureNames.Count == 0)
            return JsonResult(new
            {
                count = 0,
                items = Array.Empty<object>(),
                message = "No signatures found"
            });

        List<object> signatureList = [];
        for (var i = 0; i < signatureNames.Count; i++)
        {
            var signatureName = signatureNames[i];
            var signatureInfo = new Dictionary<string, object?>
            {
                ["index"] = i,
                ["name"] = signatureName
            };

            try
            {
                signatureInfo["isValid"] = pdfSign.VerifySignature(signatureName);
            }
            catch
            {
                signatureInfo["isValid"] = false;
            }

            try
            {
                _ = pdfSign.ExtractCertificate(signatureName);
                signatureInfo["hasCertificate"] = true;
            }
            catch
            {
                signatureInfo["hasCertificate"] = false;
            }

            signatureList.Add(signatureInfo);
        }

        return JsonResult(new
        {
            count = signatureList.Count,
            items = signatureList
        });
    }
}
