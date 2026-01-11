using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Protection;

/// <summary>
///     Handler for removing protection from Word documents.
/// </summary>
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
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var password = parameters.GetOptional<string?>("password");

        var doc = context.Document;
        var previousProtectionType = doc.ProtectionType;

        if (previousProtectionType == ProtectionType.NoProtection)
            return Success("Document is not protected, no need to unprotect");

        try
        {
            doc.Unprotect(password);
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

        return Success($"Protection removed successfully (was: {previousProtectionType})");
    }
}
