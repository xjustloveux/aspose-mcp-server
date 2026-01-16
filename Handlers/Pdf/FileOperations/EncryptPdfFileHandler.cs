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
        var encryptParams = ExtractEncryptParameters(parameters);

        var document = context.Document;
        document.Encrypt(encryptParams.UserPassword, encryptParams.OwnerPassword,
            Permissions.PrintDocument | Permissions.ModifyContent,
            CryptoAlgorithm.AESx256);

        MarkModified(context);

        return Success("PDF encrypted with password.");
    }

    /// <summary>
    ///     Extracts encrypt parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted encrypt parameters.</returns>
    private static EncryptParameters ExtractEncryptParameters(OperationParameters parameters)
    {
        return new EncryptParameters(
            parameters.GetRequired<string>("userPassword"),
            parameters.GetRequired<string>("ownerPassword")
        );
    }

    /// <summary>
    ///     Record to hold encrypt parameters.
    /// </summary>
    private record EncryptParameters(string UserPassword, string OwnerPassword);
}
