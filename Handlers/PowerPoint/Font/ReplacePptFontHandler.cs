using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Font;

/// <summary>
///     Handler for replacing a font in a PowerPoint presentation.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class ReplacePptFontHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "replace";

    /// <summary>
    ///     Replaces one font with another throughout the presentation.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: sourceFont, targetFont.
    /// </param>
    /// <returns>Success message with replacement details.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing.</exception>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var sourceFont = parameters.GetRequired<string>("sourceFont");
        var targetFont = parameters.GetRequired<string>("targetFont");

        var presentation = context.Document;
        var srcFontData = new FontData(sourceFont);
        var dstFontData = new FontData(targetFont);

        presentation.FontsManager.ReplaceFont(srcFontData, dstFontData);

        MarkModified(context);
        return new SuccessResult { Message = $"Font '{sourceFont}' replaced with '{targetFont}'." };
    }
}
