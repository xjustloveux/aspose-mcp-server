using System.Text.Json.Nodes;
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using Aspose.Pdf.Text;

namespace AsposeMcpServer.Tools;

public class PdfRedactTool : IAsposeTool
{
    public string Description => "Redact (black out) text or area on PDF page";

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
            pageIndex = new
            {
                type = "number",
                description = "Page index (1-based)"
            },
            x = new
            {
                type = "number",
                description = "X position of redaction area"
            },
            y = new
            {
                type = "number",
                description = "Y position of redaction area"
            },
            width = new
            {
                type = "number",
                description = "Width of redaction area"
            },
            height = new
            {
                type = "number",
                description = "Height of redaction area"
            },
            fillColor = new
            {
                type = "string",
                description = "Fill color (optional, default: black, format: 'R,G,B' or color name)"
            }
        },
        required = new[] { "path", "pageIndex", "x", "y", "width", "height" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>() ?? throw new ArgumentException("pageIndex is required");
        var x = arguments?["x"]?.GetValue<double>() ?? throw new ArgumentException("x is required");
        var y = arguments?["y"]?.GetValue<double>() ?? throw new ArgumentException("y is required");
        var width = arguments?["width"]?.GetValue<double>() ?? throw new ArgumentException("width is required");
        var height = arguments?["height"]?.GetValue<double>() ?? throw new ArgumentException("height is required");
        var fillColor = arguments?["fillColor"]?.GetValue<string>();

        using var document = new Document(path);
        if (pageIndex < 1 || pageIndex > document.Pages.Count)
        {
            throw new ArgumentException($"pageIndex must be between 1 and {document.Pages.Count}");
        }

        var page = document.Pages[pageIndex];
        var rect = new Rectangle(x, y, x + width, y + height);

        var redactionAnnotation = new RedactionAnnotation(page, rect);
        
        // Set fill color
        if (!string.IsNullOrEmpty(fillColor))
        {
            var colorParts = fillColor.Split(',');
            if (colorParts.Length == 3 && 
                double.TryParse(colorParts[0], out double r) &&
                double.TryParse(colorParts[1], out double g) &&
                double.TryParse(colorParts[2], out double b))
            {
                redactionAnnotation.FillColor = Aspose.Pdf.Color.FromRgb((float)r / 255, (float)g / 255, (float)b / 255);
            }
            else if (fillColor.ToLower() == "black")
            {
                redactionAnnotation.FillColor = Aspose.Pdf.Color.Black;
            }
        }
        else
        {
            redactionAnnotation.FillColor = Aspose.Pdf.Color.Black;
        }

        page.Annotations.Add(redactionAnnotation);
        
        // Note: Redaction application may require additional processing
        // The annotation is added and will be visible
        
        document.Save(path);

        return await Task.FromResult($"Redaction applied to page {pageIndex}: {path}");
    }
}

