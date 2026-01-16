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
        var protectParams = ExtractProtectParameters(parameters);

        if (string.IsNullOrWhiteSpace(protectParams.Password))
            throw new ArgumentException(
                "Password is required for protect operation. Please provide a non-empty password.");

        var doc = context.Document;
        var protectionTypeEnum = GetProtectionType(protectParams.ProtectionType);

        doc.Protect(protectionTypeEnum, protectParams.Password);

        MarkModified(context);

        return Success($"Document protected with {protectionTypeEnum}");
    }

    /// <summary>
    ///     Converts a protection type string to ProtectionType enum.
    /// </summary>
    /// <param name="protectionTypeStr">The protection type string.</param>
    /// <returns>The ProtectionType enum value.</returns>
    private static ProtectionType GetProtectionType(string protectionTypeStr)
    {
        if (Enum.TryParse<ProtectionType>(protectionTypeStr, true, out var result))
            return result;

        return ProtectionType.ReadOnly;
    }

    /// <summary>
    ///     Extracts protect parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted protect parameters.</returns>
    private static ProtectParameters ExtractProtectParameters(OperationParameters parameters)
    {
        return new ProtectParameters(
            parameters.GetOptional<string?>("password"),
            parameters.GetOptional("protectionType", "ReadOnly")
        );
    }

    /// <summary>
    ///     Record to hold protect parameters.
    /// </summary>
    /// <param name="Password">The protection password.</param>
    /// <param name="ProtectionType">The protection type.</param>
    private record ProtectParameters(string? Password, string ProtectionType);
}
