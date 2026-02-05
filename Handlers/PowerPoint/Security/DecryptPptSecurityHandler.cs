using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Security;

/// <summary>
///     Handler for removing encryption from a PowerPoint presentation.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class DecryptPptSecurityHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "decrypt";

    /// <summary>
    ///     Removes encryption from the presentation.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">No required parameters.</param>
    /// <returns>A <see cref="SuccessResult" /> indicating the encryption was removed.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var presentation = context.Document;
        presentation.ProtectionManager.RemoveEncryption();
        MarkModified(context);
        return new SuccessResult { Message = "Presentation decrypted successfully." };
    }
}
