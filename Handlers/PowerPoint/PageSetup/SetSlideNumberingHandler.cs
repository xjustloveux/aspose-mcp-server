using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.PageSetup;

/// <summary>
///     Handler for setting slide numbering in PowerPoint presentations.
/// </summary>
public class SetSlideNumberingHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "set_slide_numbering";

    /// <summary>
    ///     Sets slide numbering visibility and start number.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Optional: showSlideNumber (default: true), firstNumber (default: 1)
    /// </param>
    /// <returns>Success message with numbering information.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractSetSlideNumberingParameters(parameters);
        var presentation = context.Document;

        presentation.FirstSlideNumber = p.FirstNumber;
        presentation.HeaderFooterManager.SetAllSlideNumbersVisibility(p.ShowSlideNumber);

        foreach (var slide in presentation.Slides)
            slide.HeaderFooterManager.SetSlideNumberVisibility(p.ShowSlideNumber);

        MarkModified(context);

        var visibilityText = p.ShowSlideNumber ? "shown" : "hidden";
        return Success($"Slide numbers {visibilityText}, starting from {p.FirstNumber}.");
    }

    private static SetSlideNumberingParameters ExtractSetSlideNumberingParameters(OperationParameters parameters)
    {
        return new SetSlideNumberingParameters(
            parameters.GetOptional("showSlideNumber", true),
            parameters.GetOptional("firstNumber", 1));
    }

    private sealed record SetSlideNumberingParameters(bool ShowSlideNumber, int FirstNumber);
}
