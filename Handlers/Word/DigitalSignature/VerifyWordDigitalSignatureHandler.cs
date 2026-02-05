using Aspose.Words;
using Aspose.Words.DigitalSignatures;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Word.DigitalSignature;

namespace AsposeMcpServer.Handlers.Word.DigitalSignature;

/// <summary>
///     Handler for verifying digital signatures in a Word document.
/// </summary>
[ResultType(typeof(VerifySignaturesResult))]
public class VerifyWordDigitalSignatureHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "verify";

    /// <summary>
    ///     Verifies all digital signatures in a Word document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: path (file path to the document)
    /// </param>
    /// <returns>A result containing the verification status.</returns>
    /// <exception cref="ArgumentException">Thrown when the path parameter is missing.</exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var path = parameters.GetRequired<string>("path");
        var signatures = DigitalSignatureUtil.LoadSignatures(path);

        var totalCount = signatures.Count;
        var validCount = signatures.Count(s => s.IsValid);
        var allValid = totalCount > 0 && validCount == totalCount;

        var message = totalCount == 0
            ? "No digital signatures found in the document."
            : allValid
                ? $"All {totalCount} digital signature(s) are valid."
                : $"{validCount} of {totalCount} digital signature(s) are valid.";

        return new VerifySignaturesResult
        {
            Message = message,
            AllValid = allValid,
            TotalCount = totalCount,
            ValidCount = validCount
        };
    }
}
