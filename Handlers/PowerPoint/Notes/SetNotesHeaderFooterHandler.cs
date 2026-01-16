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
        var p = ExtractNotesHeaderFooterParameters(parameters);

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

        return Success($"Notes master header/footer updated ({string.Join(", ", settings)}).");
    }

    /// <summary>
    ///     Extracts notes header/footer parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted notes header/footer parameters.</returns>
    private static NotesHeaderFooterParameters ExtractNotesHeaderFooterParameters(OperationParameters parameters)
    {
        return new NotesHeaderFooterParameters(
            parameters.GetOptional<string?>("headerText"),
            parameters.GetOptional<string?>("footerText"),
            parameters.GetOptional<string?>("dateText"),
            parameters.GetOptional("showPageNumber", true));
    }

    /// <summary>
    ///     Record for holding notes header/footer parameters.
    /// </summary>
    /// <param name="HeaderText">The header text.</param>
    /// <param name="FooterText">The footer text.</param>
    /// <param name="DateText">The date text.</param>
    /// <param name="ShowPageNumber">Whether to show page numbers.</param>
    private sealed record NotesHeaderFooterParameters(
        string? HeaderText,
        string? FooterText,
        string? DateText,
        bool ShowPageNumber);
}
