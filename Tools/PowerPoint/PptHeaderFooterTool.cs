using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint headers and footers (set header, set footer, batch set, set slide numbering)
///     Merges: PptSetHeaderTool, PptSetFooterTool, PptBatchSetHeaderFooterTool, PptSetSlideNumberingTool
/// </summary>
public class PptHeaderFooterTool : IAsposeTool
{
    public string Description =>
        @"Manage PowerPoint headers and footers. Supports 4 operations: set_header, set_footer, batch_set, set_slide_numbering.

Usage examples:
- Set header: ppt_header_footer(operation='set_header', path='presentation.pptx', headerText='Header Text')
- Set footer: ppt_header_footer(operation='set_footer', path='presentation.pptx', footerText='Footer Text')
- Batch set: ppt_header_footer(operation='batch_set', path='presentation.pptx', headerText='Header', footerText='Footer', slideIndices=[0,1,2])
- Set slide numbering: ppt_header_footer(operation='set_slide_numbering', path='presentation.pptx', showSlideNumber=true)";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'set_header': Set header text (required params: path, headerText)
- 'set_footer': Set footer text (required params: path, footerText)
- 'batch_set': Batch set header/footer (required params: path)
- 'set_slide_numbering': Set slide numbering (required params: path)",
                @enum = new[] { "set_header", "set_footer", "batch_set", "set_slide_numbering" }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path (required for all operations)"
            },
            headerText = new
            {
                type = "string",
                description = "Header text (required for set_header)"
            },
            footerText = new
            {
                type = "string",
                description = "Footer text (optional, for set_footer/batch_set)"
            },
            slideIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description =
                    "Slide indices (0-based, optional, for set_header/batch_set, if not provided applies to all slides)"
            },
            showSlideNumber = new
            {
                type = "boolean",
                description = "Show slide number (optional, for set_footer/batch_set, default: true)"
            },
            dateText = new
            {
                type = "string",
                description = "Date/time text (optional, for set_footer/batch_set)"
            },
            firstNumber = new
            {
                type = "number",
                description = "First slide number (optional, for set_slide_numbering, default: 1)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, for all operations, defaults to input path)"
            }
        },
        required = new[] { "operation", "path" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);

        return operation.ToLower() switch
        {
            "set_header" => await SetHeaderAsync(path, outputPath, arguments),
            "set_footer" => await SetFooterAsync(path, outputPath, arguments),
            "batch_set" => await BatchSetHeaderFooterAsync(path, outputPath, arguments),
            "set_slide_numbering" => await SetSlideNumberingAsync(path, outputPath, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Sets header text for slides
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing headerText</param>
    /// <returns>Success message</returns>
    private Task<string> SetHeaderAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var headerText = ArgumentHelper.GetString(arguments, "headerText");
            var slideIndicesArray = ArgumentHelper.GetArray(arguments, "slideIndices", false);
            var slideIndices = slideIndicesArray?.Select(x => x?.GetValue<int>()).Where(x => x.HasValue)
                .Select(x => x!.Value).ToArray();

            using var presentation = new Presentation(path);
            var slides = slideIndices?.Length > 0
                ? slideIndices.Select(i => presentation.Slides[i]).ToList()
                : presentation.Slides.ToList();

            foreach (var slide in slides)
            {
                var headerFooter = slide.HeaderFooterManager;
                headerFooter.SetFooterText(headerText);
                headerFooter.SetFooterVisibility(true);
            }

            presentation.Save(outputPath, SaveFormat.Pptx);

            return $"Header set for {slides.Count} slide(s). Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets footer text for slides
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing footerText</param>
    /// <returns>Success message</returns>
    private Task<string> SetFooterAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var footerText = ArgumentHelper.GetStringNullable(arguments, "footerText");
            var showSlideNumber = ArgumentHelper.GetBool(arguments, "showSlideNumber", true);
            var dateText = ArgumentHelper.GetStringNullable(arguments, "dateText");

            using var presentation = new Presentation(path);
            foreach (var slide in presentation.Slides)
            {
                var manager = slide.HeaderFooterManager;

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

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Footer/page number settings updated. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets header and footer for multiple slides
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing headerText, footerText, optional slideIndexes</param>
    /// <returns>Success message</returns>
    private Task<string> BatchSetHeaderFooterAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var footerText = ArgumentHelper.GetStringNullable(arguments, "footerText");
            var showSlideNumber = ArgumentHelper.GetBool(arguments, "showSlideNumber", true);
            var dateText = ArgumentHelper.GetStringNullable(arguments, "dateText");
            var slideIndicesArray = ArgumentHelper.GetArray(arguments, "slideIndices", false);
            var slideIndices = slideIndicesArray?.Select(x => x?.GetValue<int>() ?? -1).ToArray();

            using var presentation = new Presentation(path);
            var targets = slideIndices?.Length > 0
                ? slideIndices
                : Enumerable.Range(0, presentation.Slides.Count).ToArray();

            foreach (var idx in targets)
                if (idx < 0 || idx >= presentation.Slides.Count)
                    throw new ArgumentException($"slide index {idx} out of range");

            foreach (var idx in targets)
            {
                var manager = presentation.Slides[idx].HeaderFooterManager;

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

            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Batch updated footer/page number/date for {targets.Length} slides. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Sets slide numbering
    /// </summary>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="arguments">JSON arguments containing isVisible, optional startNumber</param>
    /// <returns>Success message</returns>
    private Task<string> SetSlideNumberingAsync(string path, string outputPath, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var firstNumber = ArgumentHelper.GetInt(arguments, "firstNumber", 1);

            using var presentation = new Presentation(path);
            presentation.FirstSlideNumber = firstNumber;
            presentation.Save(outputPath, SaveFormat.Pptx);

            return $"Starting page number set to {firstNumber}. Output: {outputPath}";
        });
    }
}