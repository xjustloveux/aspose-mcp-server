using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.FileOperations;

/// <summary>
///     Handler for decrypting a password-protected PDF document.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class DecryptPdfFileHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "decrypt";

    /// <summary>
    ///     Decrypts a PDF document by removing password protection.
    ///     The document must already be opened with the correct password.
    /// </summary>
    /// <param name="context">The document context (document opened with password).</param>
    /// <param name="parameters">No additional parameters required.</param>
    /// <returns>Success message.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var document = context.Document;
        document.Decrypt();

        MarkModified(context);

        return new SuccessResult { Message = "PDF decrypted successfully. Password protection removed." };
    }
}
