using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.PageSetup;

/// <summary>
///     Handler for setting footer in PowerPoint presentations.
/// </summary>
public class SetFooterHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "set_footer";

    /// <summary>
    ///     Sets footer text, date, and slide number for slides.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Optional: footerText, showSlideNumber (default: true), dateText, slideIndices
    /// </param>
    /// <returns>Success message with number of slides updated.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var footerText = parameters.GetOptional<string?>("footerText");
        var showSlideNumber = parameters.GetOptional("showSlideNumber", true);
        var dateText = parameters.GetOptional<string?>("dateText");
        var slideIndices = parameters.GetOptional<int[]?>("slideIndices");

        var presentation = context.Document;
        var slides = GetTargetSlides(presentation, slideIndices);
        var applyToAll = slideIndices == null || slideIndices.Length == 0;

        if (applyToAll)
            EnableMasterVisibility(presentation, footerText, showSlideNumber, dateText);

        foreach (var slide in slides)
            ApplyFooterSettings(slide.HeaderFooterManager, footerText, showSlideNumber, dateText);

        MarkModified(context);

        return Success($"Footer settings updated for {slides.Count} slide(s).");
    }

    /// <summary>
    ///     Gets the target slides for footer settings based on provided indices.
    /// </summary>
    /// <param name="presentation">The presentation to get slides from.</param>
    /// <param name="slideIndices">The specific slide indices, or null for all slides.</param>
    /// <returns>A list of target slides.</returns>
    private static List<ISlide> GetTargetSlides(IPresentation presentation, int[]? slideIndices)
    {
        if (slideIndices == null || slideIndices.Length == 0)
            return presentation.Slides.ToList();

        List<ISlide> slides = [];
        foreach (var index in slideIndices)
        {
            PowerPointHelper.ValidateSlideIndex(index, presentation);
            slides.Add(presentation.Slides[index]);
        }

        return slides;
    }

    /// <summary>
    ///     Enables footer visibility at the master level for all slides.
    /// </summary>
    /// <param name="presentation">The presentation to configure.</param>
    /// <param name="footerText">The footer text to display.</param>
    /// <param name="showSlideNumber">Whether to show slide numbers.</param>
    /// <param name="dateText">The date text to display.</param>
    private static void EnableMasterVisibility(IPresentation presentation, string? footerText, bool showSlideNumber,
        string? dateText)
    {
        var manager = presentation.HeaderFooterManager;

        if (!string.IsNullOrEmpty(footerText))
            manager.SetAllFootersVisibility(true);

        manager.SetAllSlideNumbersVisibility(showSlideNumber);

        if (!string.IsNullOrEmpty(dateText))
            manager.SetAllDateTimesVisibility(true);
    }

    /// <summary>
    ///     Applies footer settings to a specific slide's header/footer manager.
    /// </summary>
    /// <param name="manager">The slide's header/footer manager.</param>
    /// <param name="footerText">The footer text to display.</param>
    /// <param name="showSlideNumber">Whether to show slide numbers.</param>
    /// <param name="dateText">The date text to display.</param>
    private static void ApplyFooterSettings(ISlideHeaderFooterManager manager, string? footerText,
        bool showSlideNumber, string? dateText)
    {
        if (!string.IsNullOrEmpty(footerText))
        {
            manager.SetFooterText(footerText);
            manager.SetFooterVisibility(true);
        }
        else
        {
            manager.SetFooterVisibility(false);
        }

        manager.SetSlideNumberVisibility(showSlideNumber);

        if (!string.IsNullOrEmpty(dateText))
        {
            manager.SetDateTimeText(dateText);
            manager.SetDateTimeVisibility(true);
        }
        else
        {
            manager.SetDateTimeVisibility(false);
        }
    }
}
