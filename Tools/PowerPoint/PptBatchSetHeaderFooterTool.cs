using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;
using System.Linq;

namespace AsposeMcpServer.Tools;

public class PptBatchSetHeaderFooterTool : IAsposeTool
{
    public string Description => "Batch set footer text, slide number, and date across slides";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new { type = "string", description = "Presentation file path" },
            slideIndices = new
            {
                type = "array",
                items = new { type = "number" },
                description = "Slide indices to apply (optional; default all)"
            },
            footerText = new { type = "string", description = "Footer text (optional)" },
            showSlideNumber = new { type = "boolean", description = "Show slide number (default true)" },
            dateText = new { type = "string", description = "Date/time text (optional)" }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
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
}

