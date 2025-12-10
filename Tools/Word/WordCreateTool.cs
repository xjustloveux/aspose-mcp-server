using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordCreateTool : IAsposeTool
{
    public string Description => "Create a new Word document";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Output file path"
            },
            content = new
            {
                type = "string",
                description = "Initial content (optional). Leave empty to create a blank document."
            },
            skipInitialContent = new
            {
                type = "boolean",
                description = "If true, creates a completely blank document without any initial content (default: false)"
            },
            marginTop = new
            {
                type = "number",
                description = "Top margin in points (e.g., 70.87 pt = 2.5 cm). Default: 70.87"
            },
            marginBottom = new
            {
                type = "number",
                description = "Bottom margin in points (e.g., 70.87 pt = 2.5 cm). Default: 70.87"
            },
            marginLeft = new
            {
                type = "number",
                description = "Left margin in points (e.g., 70.87 pt = 2.5 cm). Default: 70.87"
            },
            marginRight = new
            {
                type = "number",
                description = "Right margin in points (e.g., 70.87 pt = 2.5 cm). Default: 70.87"
            },
            compatibilityMode = new
            {
                type = "string",
                description = "Word compatibility mode: Word2019, Word2016, Word2013, Word2010, Word2007 (default: Word2019)",
                @enum = new[] { "Word2019", "Word2016", "Word2013", "Word2010", "Word2007" }
            },
            pageWidth = new
            {
                type = "number",
                description = "Page width in points (e.g., 595.3 for A4, 612 for Letter). Default: 595.3 (A4)"
            },
            pageHeight = new
            {
                type = "number",
                description = "Page height in points (e.g., 841.9 for A4, 792 for Letter). Default: 841.9 (A4)"
            },
            paperSize = new
            {
                type = "string",
                description = "Predefined paper size: A4, Letter, A3, Legal (default: A4). Overrides pageWidth/pageHeight if specified.",
                @enum = new[] { "A4", "Letter", "A3", "Legal" }
            },
            headerDistance = new
            {
                type = "number",
                description = "Header distance from page top in points (e.g., 45.35 pt = 1.6 cm). Default: 35.4"
            },
            footerDistance = new
            {
                type = "number",
                description = "Footer distance from page bottom in points (e.g., 45.35 pt = 1.6 cm). Default: 35.4"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var content = arguments?["content"]?.GetValue<string>();
        var skipInitialContent = arguments?["skipInitialContent"]?.GetValue<bool>() ?? false;
        var marginTop = arguments?["marginTop"]?.GetValue<double?>() ?? 70.87; // Default 2.5 cm
        var marginBottom = arguments?["marginBottom"]?.GetValue<double?>() ?? 70.87;
        var marginLeft = arguments?["marginLeft"]?.GetValue<double?>() ?? 70.87;
        var marginRight = arguments?["marginRight"]?.GetValue<double?>() ?? 70.87;
        var compatibilityMode = arguments?["compatibilityMode"]?.GetValue<string>() ?? "Word2019";
        var pageWidth = arguments?["pageWidth"]?.GetValue<double?>();
        var pageHeight = arguments?["pageHeight"]?.GetValue<double?>();
        var paperSize = arguments?["paperSize"]?.GetValue<string>() ?? "A4";
        var headerDistance = arguments?["headerDistance"]?.GetValue<double?>() ?? 35.4;
        var footerDistance = arguments?["footerDistance"]?.GetValue<double?>() ?? 35.4;

        var doc = new Document();
        
        // Set compatibility mode first (before any content modifications)
        var wordVersion = compatibilityMode switch
        {
            "Word2019" => Aspose.Words.Settings.MsWordVersion.Word2019,
            "Word2016" => Aspose.Words.Settings.MsWordVersion.Word2016,
            "Word2013" => Aspose.Words.Settings.MsWordVersion.Word2013,
            "Word2010" => Aspose.Words.Settings.MsWordVersion.Word2010,
            "Word2007" => Aspose.Words.Settings.MsWordVersion.Word2007,
            _ => Aspose.Words.Settings.MsWordVersion.Word2019
        };
        doc.CompatibilityOptions.OptimizeFor(wordVersion);
        
        // Set page setup
        var section = doc.FirstSection;
        if (section != null)
        {
            var pageSetup = section.PageSetup;
            
            // Set page size (paper size or custom dimensions)
            if (!string.IsNullOrEmpty(paperSize) && pageWidth == null && pageHeight == null)
            {
                // Use predefined paper size
                switch (paperSize.ToUpper())
                {
                    case "A4":
                        pageSetup.PageWidth = 595.3;  // 21.0 cm
                        pageSetup.PageHeight = 841.9; // 29.7 cm
                        break;
                    case "LETTER":
                        pageSetup.PageWidth = 612;    // 8.5 inch
                        pageSetup.PageHeight = 792;   // 11 inch
                        break;
                    case "A3":
                        pageSetup.PageWidth = 841.9;  // 29.7 cm
                        pageSetup.PageHeight = 1190.55; // 42.0 cm
                        break;
                    case "LEGAL":
                        pageSetup.PageWidth = 612;    // 8.5 inch
                        pageSetup.PageHeight = 1008;  // 14 inch
                        break;
                    default:
                        pageSetup.PageWidth = 595.3;  // Default to A4
                        pageSetup.PageHeight = 841.9;
                        break;
                }
            }
            else
            {
                // Use custom dimensions if specified
                if (pageWidth.HasValue)
                    pageSetup.PageWidth = pageWidth.Value;
                else
                    pageSetup.PageWidth = 595.3; // Default to A4 width
                    
                if (pageHeight.HasValue)
                    pageSetup.PageHeight = pageHeight.Value;
                else
                    pageSetup.PageHeight = 841.9; // Default to A4 height
            }
            
            // Set margins
            pageSetup.TopMargin = marginTop;
            pageSetup.BottomMargin = marginBottom;
            pageSetup.LeftMargin = marginLeft;
            pageSetup.RightMargin = marginRight;
            
            // Set header/footer distance
            pageSetup.HeaderDistance = headerDistance;
            pageSetup.FooterDistance = footerDistance;
        }
        
        var builder = new DocumentBuilder(doc);
        
        if (skipInitialContent)
        {
            // Create a completely blank document
            // Clear the default empty paragraph
            if (doc.FirstSection != null && doc.FirstSection.Body != null)
            {
                doc.FirstSection.Body.RemoveAllChildren();
                var firstPara = new Paragraph(doc);
                
                // Reset all spacing to 0 to avoid any blank lines at the top
                firstPara.ParagraphFormat.SpaceBefore = 0;
                firstPara.ParagraphFormat.SpaceAfter = 0;
                firstPara.ParagraphFormat.LineSpacing = 12; // Single line spacing
                
                doc.FirstSection.Body.AppendChild(firstPara);
                
                // Note: HeaderDistance and FooterDistance are already set in the page setup section above
                // No need to override them here
            }
        }
        else if (!string.IsNullOrEmpty(content))
        {
            // Add the provided content
            builder.Write(content);
        }
        // If content is empty/null and skipInitialContent is false, 
        // keep the default empty document structure (one empty paragraph)

        doc.Save(path);
        
        var result = $"Word document created successfully at: {path}";
        if (skipInitialContent)
            result += " (blank document)";
        else if (!string.IsNullOrEmpty(content))
            result += $" (with initial content: {content.Substring(0, Math.Min(50, content.Length))}...)";
        else
            result += " (empty document ready for content)";
            
        return await Task.FromResult(result);
    }
}

