using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Text;

namespace AsposeMcpServer.Tools;

public class PdfAddWatermarkTool : IAsposeTool
{
    public string Description => "Add text watermark to a PDF document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "PDF file path"
            },
            text = new
            {
                type = "string",
                description = "Watermark text"
            },
            opacity = new
            {
                type = "number",
                description = "Opacity (0-1, optional, default: 0.3)"
            }
        },
        required = new[] { "path", "text" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var text = arguments?["text"]?.GetValue<string>() ?? throw new ArgumentException("text is required");
        var opacity = arguments?["opacity"]?.GetValue<double>() ?? 0.3;

        using var document = new Document(path);

        foreach (Page page in document.Pages)
        {
            var watermark = new WatermarkArtifact();
            var textState = new TextState
            {
                FontSize = 72,
                ForegroundColor = Aspose.Pdf.Color.Gray,
                Font = FontRepository.FindFont("Arial")
            };

            watermark.SetTextAndState(text, textState);
            watermark.ArtifactHorizontalAlignment = HorizontalAlignment.Center;
            watermark.ArtifactVerticalAlignment = VerticalAlignment.Center;
            watermark.Rotation = 45;
            watermark.Opacity = opacity;

            page.Artifacts.Add(watermark);
        }

        document.Save(path);

        return await Task.FromResult($"Watermark added to PDF: {path}");
    }
}

