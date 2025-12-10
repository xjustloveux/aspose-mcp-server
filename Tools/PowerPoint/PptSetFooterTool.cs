using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptSetFooterTool : IAsposeTool
{
    public string Description => "Set footer text, date/time, and slide number visibility";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Presentation file path"
            },
            footerText = new
            {
                type = "string",
                description = "Footer text (optional)"
            },
            showSlideNumber = new
            {
                type = "boolean",
                description = "Show slide number (default: true)"
            },
            dateText = new
            {
                type = "string",
                description = "Date/time text (optional)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
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
}

