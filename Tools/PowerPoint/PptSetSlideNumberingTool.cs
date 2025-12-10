using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace AsposeMcpServer.Tools;

public class PptSetSlideNumberingTool : IAsposeTool
{
    public string Description => "Set the first slide number for the presentation";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new { type = "string", description = "Presentation file path" },
            firstNumber = new { type = "number", description = "First slide number (default 1)" }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var firstNumber = arguments?["firstNumber"]?.GetValue<int?>() ?? 1;

        using var presentation = new Presentation(path);
        presentation.FirstSlideNumber = firstNumber;
        presentation.Save(path, SaveFormat.Pptx);

        return await Task.FromResult($"已設定起始頁碼為 {firstNumber}");
    }
}

