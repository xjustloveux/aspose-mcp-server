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
        var p = ExtractHeaderFooterParameters(parameters);

        var presentation = context.Document;

        var handoutMaster = presentation.MasterHandoutSlideManager.MasterHandoutSlide;
        if (handoutMaster == null)
            throw new InvalidOperationException(
                "Presentation does not have a handout master slide. " +
                "Please open the presentation in PowerPoint, go to View > Handout Master to create one, then save.");

        var manager = handoutMaster.HeaderFooterManager;

        if (!string.IsNullOrEmpty(p.HeaderText))
        {
            manager.SetHeaderText(p.HeaderText);
            manager.SetHeaderVisibility(true);
        }

        if (!string.IsNullOrEmpty(p.FooterText))
        {
            manager.SetFooterText(p.FooterText);
            manager.SetFooterVisibility(true);
        }

        if (!string.IsNullOrEmpty(p.DateText))
        {
            manager.SetDateTimeText(p.DateText);
            manager.SetDateTimeVisibility(true);
        }

        manager.SetSlideNumberVisibility(p.ShowPageNumber);

        MarkModified(context);

        List<string> settings = [];
        if (!string.IsNullOrEmpty(p.HeaderText)) settings.Add("header");
        if (!string.IsNullOrEmpty(p.FooterText)) settings.Add("footer");
        if (!string.IsNullOrEmpty(p.DateText)) settings.Add("date");
        settings.Add(p.ShowPageNumber ? "page number shown" : "page number hidden");

        return Success($"Handout master header/footer updated ({string.Join(", ", settings)}).");
    }

    /// <summary>
    ///     Extracts header/footer parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted header/footer parameters.</returns>
    private static HeaderFooterParameters ExtractHeaderFooterParameters(OperationParameters parameters)
    {
        return new HeaderFooterParameters(
            parameters.GetOptional<string?>("headerText"),
            parameters.GetOptional<string?>("footerText"),
            parameters.GetOptional<string?>("dateText"),
            parameters.GetOptional("showPageNumber", true));
    }

    /// <summary>
    ///     Record for holding header/footer parameters.
    /// </summary>
    /// <param name="HeaderText">The header text.</param>
    /// <param name="FooterText">The footer text.</param>
    /// <param name="DateText">The date text.</param>
    /// <param name="ShowPageNumber">Whether to show page numbers.</param>
    private sealed record HeaderFooterParameters(
        string? HeaderText,
        string? FooterText,
        string? DateText,
        bool ShowPageNumber);
}
