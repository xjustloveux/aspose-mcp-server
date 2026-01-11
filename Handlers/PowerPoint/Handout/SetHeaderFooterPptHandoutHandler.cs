using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Handout;

/// <summary>
///     Handler for setting header and footer on PowerPoint handout master.
///     Note: Handout pages have separate header and footer fields (unlike slides which only have footer).
/// </summary>
public class SetHeaderFooterPptHandoutHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "set_header_footer";

    /// <summary>
    ///     Sets header and footer for handout master.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Optional: headerText, footerText, dateText, showPageNumber
    /// </param>
    /// <returns>Success message with updated settings.</returns>
    /// <exception cref="InvalidOperationException">
    ///     Thrown when the presentation does not have a handout master slide.
    /// </exception>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var headerText = parameters.GetOptional<string?>("headerText");
        var footerText = parameters.GetOptional<string?>("footerText");
        var dateText = parameters.GetOptional<string?>("dateText");
        var showPageNumber = parameters.GetOptional("showPageNumber", true);

        var presentation = context.Document;

        var handoutMaster = presentation.MasterHandoutSlideManager.MasterHandoutSlide;
        if (handoutMaster == null)
            throw new InvalidOperationException(
                "Presentation does not have a handout master slide. " +
                "Please open the presentation in PowerPoint, go to View > Handout Master to create one, then save.");

        var manager = handoutMaster.HeaderFooterManager;

        if (!string.IsNullOrEmpty(headerText))
        {
            manager.SetHeaderText(headerText);
            manager.SetHeaderVisibility(true);
        }

        if (!string.IsNullOrEmpty(footerText))
        {
            manager.SetFooterText(footerText);
            manager.SetFooterVisibility(true);
        }

        if (!string.IsNullOrEmpty(dateText))
        {
            manager.SetDateTimeText(dateText);
            manager.SetDateTimeVisibility(true);
        }

        manager.SetSlideNumberVisibility(showPageNumber);

        MarkModified(context);

        List<string> settings = [];
        if (!string.IsNullOrEmpty(headerText)) settings.Add("header");
        if (!string.IsNullOrEmpty(footerText)) settings.Add("footer");
        if (!string.IsNullOrEmpty(dateText)) settings.Add("date");
        settings.Add(showPageNumber ? "page number shown" : "page number hidden");

        return Success($"Handout master header/footer updated ({string.Join(", ", settings)}).");
    }
}
