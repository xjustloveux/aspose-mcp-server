using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Security;

/// <summary>
///     Handler for marking a PowerPoint presentation as final (or removing the mark).
/// </summary>
[ResultType(typeof(SuccessResult))]
public class MarkFinalPptSecurityHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "mark_final";

    /// <summary>
    ///     Marks or unmarks the presentation as final.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">Optional: markAsFinal (bool, default: true).</param>
    /// <returns>A <see cref="SuccessResult" /> indicating whether the presentation was marked or unmarked as final.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var markAsFinal = parameters.GetOptional("markAsFinal", true);
        var presentation = context.Document;
        presentation.DocumentProperties["_MarkAsFinal"] = markAsFinal;
        MarkModified(context);
        var action = markAsFinal ? "marked as final" : "unmarked as final";
        return new SuccessResult { Message = $"Presentation {action}." };
    }
}
