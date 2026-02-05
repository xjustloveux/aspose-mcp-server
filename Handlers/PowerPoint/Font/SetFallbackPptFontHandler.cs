using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Font;

/// <summary>
///     Handler for setting font fallback rules in a PowerPoint presentation.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class SetFallbackPptFontHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "set_fallback";

    /// <summary>
    ///     Sets a font fallback rule for the presentation.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: fallbackFont.
    ///     Optional: unicodeStart (default: 0x0000), unicodeEnd (default: 0xFFFF).
    /// </param>
    /// <returns>Success message with fallback rule details.</returns>
    /// <exception cref="ArgumentException">Thrown when required parameters are missing.</exception>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var fallbackFont = parameters.GetRequired<string>("fallbackFont");
        var unicodeStart = parameters.GetOptional("unicodeStart", 0x0000);
        var unicodeEnd = parameters.GetOptional("unicodeEnd", 0xFFFF);

        var presentation = context.Document;
        var rule = new FontFallBackRule((uint)unicodeStart, (uint)unicodeEnd, fallbackFont);
        presentation.FontsManager.FontFallBackRulesCollection.Add(rule);

        MarkModified(context);
        return new SuccessResult
        {
            Message =
                $"Font fallback rule set: '{fallbackFont}' for Unicode range 0x{unicodeStart:X4}-0x{unicodeEnd:X4}."
        };
    }
}
