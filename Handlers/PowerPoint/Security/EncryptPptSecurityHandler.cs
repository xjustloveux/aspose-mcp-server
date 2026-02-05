using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Security;

/// <summary>
///     Handler for encrypting a PowerPoint presentation with a password.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class EncryptPptSecurityHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "encrypt";

    /// <summary>
    ///     Encrypts the presentation with the specified password.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">Required: password (string).</param>
    /// <returns>A <see cref="SuccessResult" /> indicating the presentation was encrypted.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var password = parameters.GetRequired<string>("password");
        var presentation = context.Document;
        presentation.ProtectionManager.Encrypt(password);
        MarkModified(context);
        return new SuccessResult { Message = "Presentation encrypted successfully." };
    }
}
