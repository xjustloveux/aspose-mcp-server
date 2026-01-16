using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Layout;

/// <summary>
///     Handler for applying a theme to the presentation by copying master slides.
/// </summary>
public class ApplyThemeHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "apply_theme";

    /// <summary>
    ///     Applies a theme to the presentation by copying master slides.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: themePath
    /// </param>
    /// <returns>Success message with operation details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractApplyThemeParameters(parameters);

        if (!File.Exists(p.ThemePath))
            throw new FileNotFoundException($"Theme file not found: {p.ThemePath}");

        var presentation = context.Document;
        using var themePresentation = new Presentation(p.ThemePath);

        if (themePresentation.Masters.Count == 0)
            throw new InvalidOperationException("Theme presentation does not contain any master slides.");

        var copiedCount = 0;
        foreach (var themeMaster in themePresentation.Masters)
        {
            presentation.Masters.AddClone(themeMaster);
            copiedCount++;
        }

        if (presentation.Slides.Count > 0 && themePresentation.Masters.Count > 0)
        {
            var newMaster = presentation.Masters[^1];
            if (newMaster.LayoutSlides.Count > 0)
            {
                var defaultLayout = newMaster.LayoutSlides[0];
                foreach (var slide in presentation.Slides)
                    slide.LayoutSlide = defaultLayout;
            }
        }

        MarkModified(context);

        return Success($"Theme applied ({copiedCount} master(s) copied, layout applied to all slides).");
    }

    /// <summary>
    ///     Extracts apply theme parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted apply theme parameters.</returns>
    private static ApplyThemeParameters ExtractApplyThemeParameters(OperationParameters parameters)
    {
        return new ApplyThemeParameters(parameters.GetRequired<string>("themePath"));
    }

    /// <summary>
    ///     Record for holding apply theme parameters.
    /// </summary>
    /// <param name="ThemePath">The theme file path.</param>
    private sealed record ApplyThemeParameters(string ThemePath);
}
