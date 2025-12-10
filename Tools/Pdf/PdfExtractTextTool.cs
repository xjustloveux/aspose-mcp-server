using System.Text.Json.Nodes;
using System.Linq;
using Aspose.Pdf;
using Aspose.Pdf.Text;

namespace AsposeMcpServer.Tools;

public class PdfExtractTextTool : IAsposeTool
{
    public string Description => "Extract text from specific pages of a PDF document";

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
                description = "Page index to extract (1-based, optional - extracts all if not specified)"
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
        var pageIndex = arguments?["pageIndex"]?.GetValue<int>();
        var includeFontInfo = arguments?["includeFontInfo"]?.GetValue<bool>() ?? false;

        using var document = new Document(path);
        
        if (includeFontInfo)
        {
            // Use custom text absorber to get font information
            var sb = new System.Text.StringBuilder();
            
            var pagesToProcess = pageIndex.HasValue 
                ? new[] { pageIndex.Value - 1 } // Convert to 0-based
                : Enumerable.Range(0, document.Pages.Count).ToArray();
            
            foreach (var pageIdx in pagesToProcess)
            {
                if (pageIdx < 0 || pageIdx >= document.Pages.Count)
                    continue;
                    
                var page = document.Pages[pageIdx + 1]; // 1-based
                sb.AppendLine($"=== Page {pageIdx + 1} ===");
                
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
                        
                        sb.AppendLine($"  Position: X={textFragment.Position.XIndent}, Y={textFragment.Position.YIndent}");
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
            if (pageIndex.HasValue)
            {
                document.Pages[pageIndex.Value].Accept(textAbsorber);
            }
            else
            {
                document.Pages.Accept(textAbsorber);
            }
            return await Task.FromResult(textAbsorber.Text);
        }
    }
}

