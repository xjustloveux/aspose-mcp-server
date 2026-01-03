using System.ComponentModel;
using Aspose.Slides;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Tool for managing PowerPoint handout settings (header/footer).
///     Note: Handout pages have separate header and footer fields (unlike slides which only have footer).
/// </summary>
[McpServerToolType]
public class PptHandoutTool
{
    /// <summary>
    ///     Session manager for document lifecycle management.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptHandoutTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    public PptHandoutTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "ppt_handout")]
    [Description(@"Manage PowerPoint handout settings. Supports 1 operation: set_header_footer.

Note: Handout pages have separate header and footer fields (unlike slides which only have footer).
Important: Presentation must have a handout master (created via PowerPoint: View > Handout Master).

Usage examples:
- Set header/footer: ppt_handout(operation='set_header_footer', path='presentation.pptx', headerText='Header', footerText='Footer')")]
    public string Execute(
        [Description("Operation: set_header_footer")]
        string operation,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Header text for handout pages")]
        string? headerText = null,
        [Description("Footer text for handout pages")]
        string? footerText = null,
        [Description("Date/time text for handout pages")]
        string? dateText = null,
        [Description("Show page number on handout pages (default: true)")]
        bool showPageNumber = true)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path);

        return operation.ToLower() switch
        {
            "set_header_footer" => SetHandoutHeaderFooter(ctx, outputPath, headerText, footerText, dateText,
                showPageNumber),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Sets header and footer for handout master.
    ///     Note: Handout pages have separate header and footer fields (unlike slides which only have footer).
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="headerText">The header text.</param>
    /// <param name="footerText">The footer text.</param>
    /// <param name="dateText">The date/time text.</param>
    /// <param name="showPageNumber">Whether to show page numbers.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="InvalidOperationException">Thrown when the presentation does not have a handout master slide.</exception>
    private static string SetHandoutHeaderFooter(DocumentContext<Presentation> ctx, string? outputPath,
        string? headerText, string? footerText, string? dateText, bool showPageNumber)
    {
        var presentation = ctx.Document;

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

        ctx.Save(outputPath);

        List<string> settings = [];
        if (!string.IsNullOrEmpty(headerText)) settings.Add("header");
        if (!string.IsNullOrEmpty(footerText)) settings.Add("footer");
        if (!string.IsNullOrEmpty(dateText)) settings.Add("date");
        settings.Add(showPageNumber ? "page number shown" : "page number hidden");

        var result = $"Handout master header/footer updated ({string.Join(", ", settings)}). ";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }
}