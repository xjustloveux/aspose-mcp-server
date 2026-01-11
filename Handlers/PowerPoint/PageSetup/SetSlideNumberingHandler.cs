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
        var showSlideNumber = parameters.GetOptional("showSlideNumber", true);
        var firstNumber = parameters.GetOptional("firstNumber", 1);

        var presentation = context.Document;

        presentation.FirstSlideNumber = firstNumber;
        presentation.HeaderFooterManager.SetAllSlideNumbersVisibility(showSlideNumber);

        foreach (var slide in presentation.Slides)
            slide.HeaderFooterManager.SetSlideNumberVisibility(showSlideNumber);

        MarkModified(context);

        var visibilityText = showSlideNumber ? "shown" : "hidden";
        return Success($"Slide numbers {visibilityText}, starting from {firstNumber}.");
    }
}
