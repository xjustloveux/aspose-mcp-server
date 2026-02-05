using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Font;

/// <summary>
///     Handler for removing an embedded font from a PowerPoint presentation.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class RemoveEmbeddedPptFontHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "remove_embedded";

    /// <summary>
    ///     Removes an embedded font from the presentation.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: fontName.
    /// </param>
    /// <returns>Success message with removal details.</returns>
    /// <exception cref="ArgumentException">Thrown when the font is not embedded in the presentation.</exception>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var fontName = parameters.GetRequired<string>("fontName");

        var presentation = context.Document;
        var embeddedFonts = presentation.FontsManager.GetEmbeddedFonts();
        IFontData? targetFont = null;

        foreach (var font in embeddedFonts)
            if (font.FontName.Equals(fontName, StringComparison.OrdinalIgnoreCase))
            {
                targetFont = font;
                break;
            }

        if (targetFont == null)
            throw new ArgumentException($"Font '{fontName}' is not embedded in the presentation.");

        presentation.FontsManager.RemoveEmbeddedFont(targetFont);

        MarkModified(context);
        return new SuccessResult { Message = $"Embedded font '{fontName}' removed." };
    }
}
