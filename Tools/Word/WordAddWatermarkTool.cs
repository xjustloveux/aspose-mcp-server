using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordAddWatermarkTool : IAsposeTool
{
    public string Description => "Add text watermark to a Word document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Document file path"
            },
            text = new
            {
                type = "string",
                description = "Watermark text"
            }
        },
        required = new[] { "path", "text" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var text = arguments?["text"]?.GetValue<string>() ?? throw new ArgumentException("text is required");

        var doc = new Document(path);
        
        var watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 72,
            IsSemitrasparent = true,
            Layout = WatermarkLayout.Diagonal
        };

        doc.Watermark.SetText(text, watermarkOptions);
        doc.Save(path);

        return await Task.FromResult($"Watermark added to document: {path}");
    }
}

