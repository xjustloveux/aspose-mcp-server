using System.Text.Json.Nodes;
using System.Linq;
using Aspose.Slides;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools;

/// <summary>
/// Unified tool for managing PowerPoint headers and footers (set header, set footer, batch set, set slide numbering)
/// Merges: PptSetHeaderTool, PptSetFooterTool, PptBatchSetHeaderFooterTool, PptSetSlideNumberingTool
/// </summary>
public class PptHeaderFooterTool : IAsposeTool
{
    public string Description => @"Manage PowerPoint headers and footers. Supports 4 operations: set_header, set_footer, batch_set, set_slide_numbering.

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
                description = "Slide indices (0-based, optional, for set_header/batch_set, if not provided applies to all slides)"
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
            }
        },
        required = new[] { "operation", "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = arguments?["operation"]?.GetValue<string>() ?? throw new ArgumentException("operation is required");
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");

        return operation.ToLower() switch
        {
            "set_header" => await SetHeaderAsync(arguments, path),
            "set_footer" => await SetFooterAsync(arguments, path),
            "batch_set" => await BatchSetHeaderFooterAsync(arguments, path),
            "set_slide_numbering" => await SetSlideNumberingAsync(arguments, path),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    private async Task<string> SetHeaderAsync(JsonObject? arguments, string path)
    {
        var headerText = arguments?["headerText"]?.GetValue<string>() ?? throw new ArgumentException("headerText is required for set_header operation");
        var slideIndices = arguments?["slideIndices"]?.AsArray()?.Select(x => x?.GetValue<int>()).Where(x => x.HasValue).Select(x => x!.Value).ToArray();

        using var presentation = new Presentation(path);
        var slides = slideIndices?.Length > 0
            ? slideIndices.Select(i => presentation.Slides[i]).ToList()
            : presentation.Slides.Cast<ISlide>().ToList();

        foreach (var slide in slides)
        {
            var headerFooter = slide.HeaderFooterManager;
            headerFooter.SetFooterText(headerText);
            headerFooter.SetFooterVisibility(true);
        }

        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"Header set for {slides.Count} slide(s): {path}");
    }

    private async Task<string> SetFooterAsync(JsonObject? arguments, string path)
    {
        var footerText = arguments?["footerText"]?.GetValue<string>();
        var showSlideNumber = arguments?["showSlideNumber"]?.GetValue<bool?>() ?? true;
        var dateText = arguments?["dateText"]?.GetValue<string>();

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

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult("已更新頁尾/頁碼設定");
    }

    private async Task<string> BatchSetHeaderFooterAsync(JsonObject? arguments, string path)
    {
        var footerText = arguments?["footerText"]?.GetValue<string>();
        var showSlideNumber = arguments?["showSlideNumber"]?.GetValue<bool?>() ?? true;
        var dateText = arguments?["dateText"]?.GetValue<string>();
        var slideIndices = arguments?["slideIndices"]?.AsArray()?.Select(x => x?.GetValue<int>() ?? -1).ToArray();

        using var presentation = new Presentation(path);
        var targets = slideIndices?.Length > 0
            ? slideIndices
            : Enumerable.Range(0, presentation.Slides.Count).ToArray();

        foreach (var idx in targets)
        {
            if (idx < 0 || idx >= presentation.Slides.Count)
            {
                throw new ArgumentException($"slide index {idx} out of range");
            }
        }

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

        presentation.Save(path, SaveFormat.Pptx);
        return await Task.FromResult($"已批次更新 {targets.Length} 張投影片的頁尾/頁碼/日期");
    }

    private async Task<string> SetSlideNumberingAsync(JsonObject? arguments, string path)
    {
        var firstNumber = arguments?["firstNumber"]?.GetValue<int?>() ?? 1;

        using var presentation = new Presentation(path);
        presentation.FirstSlideNumber = firstNumber;
        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"已設定起始頁碼為 {firstNumber}");
    }
}

