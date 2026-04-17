using Aspose.Pdf;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Pdf.FileOperations;

/// <summary>
///     Handler for encrypting a PDF document with passwords, an optional encryption algorithm,
///     optional permission flags, and an optional PDF 2.0 mode switch.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class EncryptPdfFileHandler : OperationHandlerBase<Document>
{
    /// <summary>
    ///     All valid <see cref="CryptoAlgorithm" /> names exposed to callers.
    /// </summary>
    private static readonly string[] ValidAlgorithms = Enum.GetNames<CryptoAlgorithm>();

    /// <summary>
    ///     All valid <see cref="Permissions" /> flag names exposed to callers.
    /// </summary>
    private static readonly string[] ValidPermissions = Enum.GetNames<Permissions>();

    /// <inheritdoc />
    public override string Operation => "encrypt";

    /// <summary>
    ///     Encrypts a PDF document with user and owner passwords, an optional encryption algorithm,
    ///     optional document permissions, and an optional PDF 2.0 flag.
    /// </summary>
    /// <param name="context">The document context holding the in-memory <see cref="Document" />.</param>
    /// <param name="parameters">
    ///     Required: userPassword, ownerPassword.
    ///     Optional: algorithm (default AESx256), permissions (default PrintDocument|ModifyContent),
    ///     usePdf20 (default false).
    /// </param>
    /// <returns>A <see cref="SuccessResult" /> with a confirmation message.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when algorithm is unrecognised, a permissions entry is invalid,
    ///     or usePdf20=true is combined with an algorithm other than AESx256.
    /// </exception>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var encryptParams = ExtractEncryptParameters(parameters);

        var algorithm = ParseAlgorithm(encryptParams.Algorithm);
        var permissions = ParsePermissions(encryptParams.Permissions);

        if (encryptParams.UsePdf20 && algorithm != CryptoAlgorithm.AESx256)
            throw new ArgumentException("usePdf20=true requires algorithm=AESx256.");

        var document = context.Document;

        if (encryptParams.UsePdf20)
            document.Encrypt(encryptParams.UserPassword, encryptParams.OwnerPassword, permissions, algorithm, true);
        else
            document.Encrypt(encryptParams.UserPassword, encryptParams.OwnerPassword, permissions, algorithm);

        MarkModified(context);

        return new SuccessResult { Message = "PDF encrypted with password." };
    }

    /// <summary>
    ///     Parses an optional algorithm name into a <see cref="CryptoAlgorithm" /> value.
    /// </summary>
    /// <param name="algorithm">
    ///     A case-insensitive algorithm name (e.g. "AESx256").
    ///     Pass <see langword="null" /> or an empty string to accept the default (<see cref="CryptoAlgorithm.AESx256" />).
    /// </param>
    /// <returns>The resolved <see cref="CryptoAlgorithm" />.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when the value is not null/empty and does not match a valid enum member.
    ///     The message lists all accepted values.
    /// </exception>
    private static CryptoAlgorithm ParseAlgorithm(string? algorithm)
    {
        if (string.IsNullOrEmpty(algorithm))
            return CryptoAlgorithm.AESx256;

        if (!Enum.TryParse<CryptoAlgorithm>(algorithm, true, out var result))
            throw new ArgumentException(
                $"Invalid algorithm '{algorithm}'. Valid values: {string.Join(", ", ValidAlgorithms)}.");

        return result;
    }

    /// <summary>
    ///     Parses an optional array of permission names into a combined <see cref="Permissions" /> flags value.
    /// </summary>
    /// <param name="permissionNames">
    ///     A case-insensitive array of permission flag names (e.g. ["PrintDocument", "ModifyContent"]).
    ///     Pass <see langword="null" /> to accept the default (PrintDocument | ModifyContent).
    ///     Pass an empty array to grant no permissions (equivalent to <c>(Permissions)0</c>).
    /// </param>
    /// <returns>The combined <see cref="Permissions" /> flags value.</returns>
    /// <exception cref="ArgumentException">
    ///     Thrown when any entry in the array does not match a valid <see cref="Permissions" /> enum member.
    ///     The message lists all accepted values.
    /// </exception>
    private static Permissions ParsePermissions(string[]? permissionNames)
    {
        if (permissionNames == null)
            return Permissions.PrintDocument | Permissions.ModifyContent;

        if (permissionNames.Length == 0)
            return 0;

        var combined = (Permissions)0;
        foreach (var name in permissionNames)
        {
            if (!Enum.TryParse<Permissions>(name, true, out var flag))
                throw new ArgumentException(
                    $"Invalid permission '{name}'. Valid values: {string.Join(", ", ValidPermissions)}.");

            combined |= flag;
        }

        return combined;
    }

    /// <summary>
    ///     Extracts encrypt parameters from the operation parameters bag.
    /// </summary>
    /// <param name="parameters">
    ///     Must contain "userPassword" and "ownerPassword" as non-null strings.
    ///     May optionally contain "algorithm" (string), "permissions" (string[]), and "usePdf20" (bool).
    /// </param>
    /// <returns>An <see cref="EncryptParameters" /> record populated from the parameters bag.</returns>
    /// <exception cref="ArgumentException">Thrown when userPassword or ownerPassword is missing.</exception>
    private static EncryptParameters ExtractEncryptParameters(OperationParameters parameters)
    {
        return new EncryptParameters(
            parameters.GetRequired<string>("userPassword"),
            parameters.GetRequired<string>("ownerPassword"),
            parameters.GetOptional<string?>("algorithm"),
            parameters.GetOptional<string[]?>("permissions"),
            parameters.GetOptional<bool>("usePdf20")
        );
    }

    /// <summary>
    ///     Holds all parameters required to perform a PDF encryption operation.
    /// </summary>
    /// <param name="UserPassword">Password required to open the encrypted PDF.</param>
    /// <param name="OwnerPassword">Password that grants full permissions over the encrypted PDF.</param>
    /// <param name="Algorithm">
    ///     Optional algorithm name; <see langword="null" /> resolves to <see cref="CryptoAlgorithm.AESx256" />.
    /// </param>
    /// <param name="Permissions">
    ///     Optional array of permission flag names; <see langword="null" /> resolves to
    ///     PrintDocument | ModifyContent. An empty array resolves to no permissions (<c>(Permissions)0</c>).
    /// </param>
    /// <param name="UsePdf20">
    ///     When <see langword="true" />, encrypts the document using the PDF 2.0 security revision.
    ///     Requires <paramref name="Algorithm" /> to resolve to <see cref="CryptoAlgorithm.AESx256" />.
    /// </param>
    private sealed record EncryptParameters(
        string UserPassword,
        string OwnerPassword,
        string? Algorithm,
        string[]? Permissions,
        bool UsePdf20);
}
