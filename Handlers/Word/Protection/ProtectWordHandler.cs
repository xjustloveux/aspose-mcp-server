using Aspose.Words;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.Word.Protection;

/// <summary>
///     Handler for protecting Word documents.
/// </summary>
public class ProtectWordHandler : OperationHandlerBase<Document>
{
    /// <inheritdoc />
    public override string Operation => "protect";

    /// <summary>
    ///     Protects the document with specified protection type and password.
    /// </summary>
    /// <param name="context">The document context.</param>
    /// <param name="parameters">
    ///     Required: password
    ///     Optional: protectionType (default: ReadOnly)
    /// </param>
    /// <returns>Success message with protection details.</returns>
    public override string Execute(OperationContext<Document> context, OperationParameters parameters)
    {
        var password = parameters.GetOptional<string?>("password");
        var protectionType = parameters.GetOptional("protectionType", "ReadOnly");

        if (string.IsNullOrWhiteSpace(password))
            throw new ArgumentException(
                "Password is required for protect operation. Please provide a non-empty password.");

        var doc = context.Document;
        var protectionTypeEnum = GetProtectionType(protectionType);

        doc.Protect(protectionTypeEnum, password);

        MarkModified(context);

        return Success($"Document protected with {protectionTypeEnum}");
    }

    /// <summary>
    ///     Converts a protection type string to ProtectionType enum.
    /// </summary>
    private static ProtectionType GetProtectionType(string protectionTypeStr)
    {
        if (Enum.TryParse<ProtectionType>(protectionTypeStr, true, out var result))
            return result;

        return ProtectionType.ReadOnly;
    }
}
