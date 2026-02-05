using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Results.PowerPoint.Font;

namespace AsposeMcpServer.Handlers.PowerPoint.Font;

/// <summary>
///     Handler for getting used fonts in a PowerPoint presentation.
/// </summary>
[ResultType(typeof(GetFontsPptResult))]
public class GetUsedPptFontsHandler : OperationHandlerBase<Presentation>
{
    /// <summary>
    ///     Common system font names used for the IsCustom heuristic.
    /// </summary>
    private static readonly HashSet<string> SystemFonts = new(StringComparer.OrdinalIgnoreCase)
    {
        "Arial", "Calibri", "Cambria", "Comic Sans MS", "Consolas",
        "Courier New", "Georgia", "Impact", "Lucida Console",
        "Segoe UI", "Tahoma", "Times New Roman", "Trebuchet MS", "Verdana",
        "Wingdings", "Symbol", "MS Gothic", "MS Mincho", "Microsoft YaHei",
        "SimSun", "SimHei", "KaiTi", "DFKai-SB", "MingLiU", "PMingLiU"
    };

    /// <inheritdoc />
    public override string Operation => "get_used";

    /// <summary>
    ///     Gets all fonts used in the presentation.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">No required parameters.</param>
    /// <returns>GetFontsPptResult containing font information.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        _ = parameters;

        var presentation = context.Document;
        var allFonts = presentation.FontsManager.GetFonts();
        var embeddedFonts = presentation.FontsManager.GetEmbeddedFonts();
        var embeddedSet =
            new HashSet<string>(embeddedFonts.Select(f => f.FontName), StringComparer.OrdinalIgnoreCase);

        var fontInfos = new List<PptFontInfo>();
        foreach (var font in allFonts)
            fontInfos.Add(new PptFontInfo
            {
                FontName = font.FontName,
                IsEmbedded = embeddedSet.Contains(font.FontName),
                IsCustom = !IsSystemFont(font.FontName)
            });

        var embeddedCount = fontInfos.Count(f => f.IsEmbedded);

        return new GetFontsPptResult
        {
            Count = fontInfos.Count,
            EmbeddedCount = embeddedCount,
            Items = fontInfos,
            Message = $"Found {fontInfos.Count} font(s), {embeddedCount} embedded."
        };
    }

    /// <summary>
    ///     Simple heuristic to check if a font is likely a system font.
    /// </summary>
    /// <param name="fontName">The font name to check.</param>
    /// <returns>True if likely a system font; otherwise, false.</returns>
    private static bool IsSystemFont(string fontName)
    {
        return SystemFonts.Contains(fontName);
    }
}
