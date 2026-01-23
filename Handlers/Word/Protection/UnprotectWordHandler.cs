using Aspose.Words;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.Word.Protection;

/// <summary>
///     Handler for removing protection from Word documents.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class UnprotectWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "unprotect";

    /// <summary>
    ///     Removes protection from the document.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Optional: password
    /// </param>
    /// <returns>Success message with unprotection details.</returns>
    public override object Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var unprotectParams = ExtractUnprotectParameters(parameters);

        var doc = context.Document;
        var previousProtectionType = doc.ProtectionType;

        if (previousProtectionType == ProtectionType.NoProtection)
            return new SuccessResult { Message = "Document is not protected, no need to unprotect" };

        try
        {
            doc.Unprotect(unprotectParams.Password);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                $"Failed to unprotect document: The password may be incorrect or the document has additional restrictions. Details: {ex.Message}",
                ex);
        }

        if (doc.ProtectionType != ProtectionType.NoProtection)
            throw new InvalidOperationException(
                "Failed to unprotect document: The password may be incorrect. Please verify the password and try again.");

        MarkModified(context);

        return new SuccessResult { Message = $"Protection removed successfully (was: {previousProtectionType})" };
    }

    /// <summary>
    ///     Extracts unprotect parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted unprotect parameters.</returns>
    private static UnprotectParameters ExtractUnprotectParameters(OperationParameters parameters)
    {
        return new UnprotectParameters(
            parameters.GetOptional<string?>("password")
        );
    }

    /// <summary>
    ///     Record to hold unprotect parameters.
    /// </summary>
    /// <param name="Password">The protection password.</param>
    private sealed record UnprotectParameters(string? Password);
}
