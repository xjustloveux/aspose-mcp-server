using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Tool for managing PowerPoint handout settings (header/footer).
///     Note: Handout pages have separate header and footer fields (unlike slides which only have footer).
/// </summary>
public class PptHandoutTool : IAsposeTool
{
    public string Description => @"Manage PowerPoint handout settings. Supports 1 operation: set_header_footer.

Note: Handout pages have separate header and footer fields (unlike slides which only have footer).
Important: Presentation must have a handout master (created via PowerPoint: View > Handout Master).

Usage examples:
- Set header/footer: ppt_handout(operation='set_header_footer', path='presentation.pptx', headerText='Header', footerText='Footer')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'set_header_footer': Set header/footer for handout master (required params: path)",
                @enum = new[] { "set_header_footer" }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path (required for all operations)"
            },
            headerText = new
            {
                type = "string",
                description = "Header text for handout pages (optional)"
            },
            footerText = new
            {
                type = "string",
                description = "Footer text for handout pages (optional)"
            },
            dateText = new
            {
                type = "string",
                description = "Date/time text for handout pages (optional)"
            },
            showPageNumber = new
            {
                type = "boolean",
                description = "Show page number on handout pages (optional, default: true)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to input path)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments.
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters.</param>
    /// <returns>Result message as a string.</returns>
    /// <exception cref="ArgumentException">Thrown when operation is unknown.</exception>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        return operation.ToLower() switch
        {
            "set_header_footer" => await SetHandoutHeaderFooterAsync(path, outputPath, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Sets header and footer for handout master.
    ///     Note: Handout pages have separate header and footer fields (unlike slides which only have footer).
    /// </summary>
    /// <param name="path">PowerPoint file path.</param>
    /// <param name="outputPath">Output file path.</param>
    /// <param name="arguments">JSON arguments containing headerText, footerText, dateText, showPageNumber.</param>
    /// <returns>Success message.</returns>
    /// <exception cref="InvalidOperationException">Thrown when presentation does not have a handout master slide.</exception>
    private Task<string> SetHandoutHeaderFooterAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var headerText = ArgumentHelper.GetStringNullable(arguments, "headerText");
            var footerText = ArgumentHelper.GetStringNullable(arguments, "footerText");
            var dateText = ArgumentHelper.GetStringNullable(arguments, "dateText");
            var showPageNumber = ArgumentHelper.GetBool(arguments, "showPageNumber", true);

            using var presentation = new Presentation(path);

            // Check if handout master exists
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

            presentation.Save(outputPath, SaveFormat.Pptx);

            var settings = new List<string>();
            if (!string.IsNullOrEmpty(headerText)) settings.Add("header");
            if (!string.IsNullOrEmpty(footerText)) settings.Add("footer");
            if (!string.IsNullOrEmpty(dateText)) settings.Add("date");
            settings.Add(showPageNumber ? "page number shown" : "page number hidden");

            return $"Handout master header/footer updated ({string.Join(", ", settings)}). Output: {outputPath}";
        });
    }
}