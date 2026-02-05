using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Security;

/// <summary>
///     Handler for setting write protection on a PowerPoint presentation.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SetWriteProtectionPptSecurityHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "set_write_protection";

    /// <summary>
    ///     Sets write protection on the presentation with the specified password.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">Required: password (string).</param>
    /// <returns>A <see cref="SuccessResult" /> indicating write protection was set.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var password = parameters.GetRequired<string>("password");
        var presentation = context.Document;
        presentation.ProtectionManager.SetWriteProtection(password);
        MarkModified(context);
        return new SuccessResult { Message = "Write protection set on presentation." };
    }
}
