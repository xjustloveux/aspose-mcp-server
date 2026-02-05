using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.Common;

namespace AsposeMcpServer.Handlers.PowerPoint.Font;

/// <summary>
///     Handler for embedding a font in a PowerPoint presentation.
/// </summary>
[ResultType(typeof(SuccessResult))]
public class EmbedPptFontHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "embed";

    /// <summary>
    ///     Embeds a font in the presentation.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: fontName.
    ///     Optional: embedMode (default: "all"). Values: "all", "subset" (only used characters).
    /// </param>
    /// <returns>Success message with embed details.</returns>
    /// <exception cref="ArgumentException">Thrown when the font is not found in the presentation.</exception>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var fontName = parameters.GetRequired<string>("fontName");
        var embedMode = parameters.GetOptional("embedMode", "all");

        var presentation = context.Document;
        var allFonts = presentation.FontsManager.GetFonts();
        IFontData? targetFont = null;

        foreach (var font in allFonts)
            if (font.FontName.Equals(fontName, StringComparison.OrdinalIgnoreCase))
            {
                targetFont = font;
                break;
            }

        if (targetFont == null)
            throw new ArgumentException(
                $"Font '{fontName}' not found in presentation. Use get_used to list available fonts.");

        var embeddedFonts = presentation.FontsManager.GetEmbeddedFonts();
        foreach (var embedded in embeddedFonts)
            if (embedded.FontName.Equals(fontName, StringComparison.OrdinalIgnoreCase))
                return new SuccessResult { Message = $"Font '{fontName}' is already embedded." };

        var rule = embedMode.ToLowerInvariant() == "subset"
            ? EmbedFontCharacters.OnlyUsed
            : EmbedFontCharacters.All;

        presentation.FontsManager.AddEmbeddedFont(targetFont, rule);

        MarkModified(context);
        return new SuccessResult { Message = $"Font '{fontName}' embedded ({embedMode})." };
    }
}
