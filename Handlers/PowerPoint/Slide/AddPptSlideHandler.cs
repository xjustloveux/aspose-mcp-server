using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Slide;

/// <summary>
///     Handler for adding slides to PowerPoint presentations.
/// </summary>
public class AddPptSlideHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "add";

    /// <summary>
    ///     Adds a new slide to the presentation.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Optional: layoutType (Blank, Title, TitleOnly, TwoColumn, SectionHeader)
    /// </param>
    /// <returns>Success message with slide count.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractAddPptSlideParameters(parameters);

        var presentation = context.Document;

        if (presentation.LayoutSlides.Count == 0)
            throw new InvalidOperationException("Presentation has no layout slides");

        var layoutType = p.LayoutType.ToLower() switch
        {
            "title" => SlideLayoutType.Title,
            "titleonly" => SlideLayoutType.TitleOnly,
            "blank" => SlideLayoutType.Blank,
            "twocolumn" => SlideLayoutType.TwoColumnText,
            "sectionheader" => SlideLayoutType.SectionHeader,
            _ => SlideLayoutType.Custom
        };

        var layoutSlide = presentation.LayoutSlides.FirstOrDefault(ls => ls.LayoutType == layoutType) ??
                          presentation.LayoutSlides[0];
        _ = presentation.Slides.AddEmptySlide(layoutSlide);

        MarkModified(context);

        return Success($"Slide added (total: {presentation.Slides.Count}).");
    }

    /// <summary>
    ///     Extracts parameters for add slide operation.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted parameters.</returns>
    private static AddPptSlideParameters ExtractAddPptSlideParameters(OperationParameters parameters)
    {
        return new AddPptSlideParameters(parameters.GetOptional("layoutType", "Blank"));
    }

    /// <summary>
    ///     Parameters for add slide operation.
    /// </summary>
    /// <param name="LayoutType">The layout type for the new slide.</param>
    private sealed record AddPptSlideParameters(string LayoutType);
}
