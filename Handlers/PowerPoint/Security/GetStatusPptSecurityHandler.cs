using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.PowerPoint.Security;

namespace AsposeMcpServer.Handlers.PowerPoint.Security;

/// <summary>
///     Handler for getting the security status of a PowerPoint presentation.
/// </summary>
[ResultType(typeof(SecurityStatusPptResult))]
public class GetStatusPptSecurityHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "get_status";

    /// <summary>
    ///     Gets the security status of the presentation including encryption, write protection,
    ///     mark-as-final, and read-only recommendation states.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">No required parameters.</param>
    /// <returns>A <see cref="SecurityStatusPptResult" /> containing the security status details.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var presentation = context.Document;
        var isEncrypted = presentation.ProtectionManager.IsEncrypted;
        var isWriteProtected = presentation.ProtectionManager.IsWriteProtected;
        var isReadOnlyRecommended = presentation.ProtectionManager.ReadOnlyRecommended;
        var markAsFinalObj = presentation.DocumentProperties["_MarkAsFinal"];
        var isMarkedFinal = markAsFinalObj is true;

        return new SecurityStatusPptResult
        {
            IsEncrypted = isEncrypted,
            IsWriteProtected = isWriteProtected,
            IsMarkedFinal = isMarkedFinal,
            IsReadOnlyRecommended = isReadOnlyRecommended,
            Message =
                $"Encrypted: {isEncrypted}, WriteProtected: {isWriteProtected}, MarkedFinal: {isMarkedFinal}, ReadOnlyRecommended: {isReadOnlyRecommended}."
        };
    }
}
