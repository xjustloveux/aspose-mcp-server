using System.Security.Cryptography.X509Certificates;
using Aspose.Words;
using Aspose.Words.DigitalSignatures;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Word.DigitalSignature;

namespace AsposeMcpServer.Handlers.Word.DigitalSignature;

/// <summary>
///     Handler for listing digital signatures in a Word document.
/// </summary>
[ResultType(typeof(GetSignaturesResult))]
public class ListWordDigitalSignatureHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "list";

    /// <summary>
    ///     Lists all digital signatures in a Word document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: path (file path to the document)
    /// </param>
    /// <returns>A result containing the list of digital signatures.</returns>
    /// <exception cref="ArgumentException">Thrown when the path parameter is missing.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        var signatures = DigitalSignatureUtil.LoadSignatures(path);

        var signatureInfos = signatures.Select(sig => new SignatureInfo
        {
            SignerName = sig.CertificateHolder?.Certificate?.GetNameInfo(
                X509NameType.SimpleName, false),
            Comments = sig.Comments,
            SignTime = sig.SignTime.ToString("o"),
            IsValid = sig.IsValid,
            IssuerName = sig.CertificateHolder?.Certificate?.Issuer,
            SubjectName = sig.CertificateHolder?.Certificate?.Subject
        }).ToList();

        return new GetSignaturesResult
        {
            Message = $"Found {signatureInfos.Count} digital signature(s).",
            Count = signatureInfos.Count,
            Signatures = signatureInfos
        };
    }
}
