using System.Text.Json.Nodes;
using System.Text;
using Aspose.Pdf;
using Aspose.Pdf.Text;

namespace AsposeMcpServer.Tools;

public class PdfGetContentTool : IAsposeTool
{
    public string Description => "Get text content from a PDF document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Input file path"
            },
            includeFontInfo = new
            {
                type = "boolean",
                description = "Include font information for each text fragment (optional, default: false)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var includeFontInfo = arguments?["includeFontInfo"]?.GetValue<bool>() ?? false;

        using var document = new Document(path);
        
        if (includeFontInfo)
        {
            // Extract text with font information
            var sb = new System.Text.StringBuilder();
            
            for (int pageIdx = 1; pageIdx <= document.Pages.Count; pageIdx++)
            {
                var page = document.Pages[pageIdx];
                sb.AppendLine($"=== Page {pageIdx} ===");
                
                foreach (var paragraph in page.Paragraphs)
                {
                    if (paragraph is TextFragment textFragment)
                    {
                        sb.AppendLine($"Text: {textFragment.Text}");
                        sb.AppendLine($"  Font: {textFragment.TextState.Font?.FontName ?? "(default)"}");
                        sb.AppendLine($"  Font Size: {textFragment.TextState.FontSize}");
                        sb.AppendLine($"  Font Style: {textFragment.TextState.FontStyle}");
                        
                        if (textFragment.TextState.ForegroundColor != null)
                        {
                            var color = textFragment.TextState.ForegroundColor;
                            sb.AppendLine($"  Color: {color}");
                        }
                        
                        sb.AppendLine();
                    }
                }
            }
            
            return await Task.FromResult(sb.ToString());
        }
        else
        {
            // Simple text extraction
            var textAbsorber = new TextAbsorber();
            document.Pages.Accept(textAbsorber);
            return await Task.FromResult(textAbsorber.Text);
        }
    }
}

