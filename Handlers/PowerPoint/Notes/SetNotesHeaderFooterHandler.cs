using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;

namespace AsposeMcpServer.Handlers.PowerPoint.Notes;

/// <summary>
///     Handler for setting header and footer on notes master.
/// </summary>
public class SetNotesHeaderFooterHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "set_header_footer";

    /// <summary>
    ///     Sets header and footer for notes master.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Optional: headerText, footerText, dateText, showPageNumber (default: true)
    /// </param>
    /// <returns>Success message indicating what settings were updated.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var headerText = parameters.GetOptional<string?>("headerText");
        var footerText = parameters.GetOptional<string?>("footerText");
        var dateText = parameters.GetOptional<string?>("dateText");
        var showPageNumber = parameters.GetOptional("showPageNumber", true);

        var presentation = context.Document;

        if (presentation.Slides.Count == 0)
            throw new InvalidOperationException(
                "Presentation has no slides. Cannot set notes header/footer on empty presentation.");

        if (presentation.MasterNotesSlideManager.MasterNotesSlide == null)
            presentation.Slides[0].NotesSlideManager.AddNotesSlide();

        var notesMaster = presentation.MasterNotesSlideManager.MasterNotesSlide;
        if (notesMaster == null)
            throw new InvalidOperationException("Failed to create notes master slide.");

        var manager = notesMaster.HeaderFooterManager;

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

        return Success($"Notes master header/footer updated ({string.Join(", ", settings)}).");
    }
}
