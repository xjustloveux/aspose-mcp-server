using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Security;

/// <summary>
///     Handler for removing write protection from a PowerPoint presentation.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class RemoveWriteProtectionPptSecurityHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "remove_write_protection";

    /// <summary>
    ///     Removes write protection from the presentation.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">No required parameters.</param>
    /// <returns>A <see cref="SuccessResult" /> indicating write protection was removed.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var presentation = context.Document;
        presentation.ProtectionManager.RemoveWriteProtection();
        MarkModified(context);
        return new SuccessResult { Message = "Write protection removed from presentation." };
    }
}
