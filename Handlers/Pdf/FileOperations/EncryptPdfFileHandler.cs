using Aspose.Pdf;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Pdf.FileOperations;

/// <summary>
///     Handler for encrypting a PDF document with passwords.
/// </summary>
public class EncryptPdfFileHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "encrypt";

    /// <summary>
    ///     Encrypts a PDF document with user and owner passwords.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: userPassword, ownerPassword
    /// </param>
    /// <returns>Success message.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var userPassword = parameters.GetRequired<string>("userPassword");
        var ownerPassword = parameters.GetRequired<string>("ownerPassword");

        var document = context.Document;
        document.Encrypt(userPassword, ownerPassword, Permissions.PrintDocument | Permissions.ModifyContent,
            CryptoAlgorithm.AESx256);

        MarkModified(context);

        return Success("PDF encrypted with password.");
    }
}
