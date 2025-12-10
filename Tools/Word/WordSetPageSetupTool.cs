using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordSetPageSetupTool : IAsposeTool
{
    public string Description => "Set page setup (margins, size, orientation) for a Word document";

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
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input)"
            },
            marginTop = new
            {
                type = "number",
                description = "Top margin in points (e.g., 70.87 pt = 2.5 cm)"
            },
            marginBottom = new
            {
                type = "number",
                description = "Bottom margin in points (e.g., 70.87 pt = 2.5 cm)"
            },
            marginLeft = new
            {
                type = "number",
                description = "Left margin in points (e.g., 70.87 pt = 2.5 cm)"
            },
            marginRight = new
            {
                type = "number",
                description = "Right margin in points (e.g., 70.87 pt = 2.5 cm)"
            },
            pageWidth = new
            {
                type = "number",
                description = "Page width in points (e.g., 595.3 pt for A4 width)"
            },
            pageHeight = new
            {
                type = "number",
                description = "Page height in points (e.g., 842 pt for A4 height)"
            },
            orientation = new
            {
                type = "string",
                description = "Page orientation: portrait, landscape",
                @enum = new[] { "portrait", "landscape" }
            },
            headerDistance = new
            {
                type = "number",
                description = "Header distance from page top in points (e.g., 45.35 pt = 1.6 cm)"
            },
            footerDistance = new
            {
                type = "number",
                description = "Footer distance from page bottom in points (e.g., 45.35 pt = 1.6 cm)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        
        var marginTop = arguments?["marginTop"]?.GetValue<double?>();
        var marginBottom = arguments?["marginBottom"]?.GetValue<double?>();
        var marginLeft = arguments?["marginLeft"]?.GetValue<double?>();
        var marginRight = arguments?["marginRight"]?.GetValue<double?>();
        var pageWidth = arguments?["pageWidth"]?.GetValue<double?>();
        var pageHeight = arguments?["pageHeight"]?.GetValue<double?>();
        var orientation = arguments?["orientation"]?.GetValue<string>();
        var headerDistance = arguments?["headerDistance"]?.GetValue<double?>();
        var footerDistance = arguments?["footerDistance"]?.GetValue<double?>();

        var doc = new Document(path);
        
        // Apply to all sections
        foreach (Section section in doc.Sections)
        {
            var pageSetup = section.PageSetup;
            
            if (marginTop.HasValue)
                pageSetup.TopMargin = marginTop.Value;
            if (marginBottom.HasValue)
                pageSetup.BottomMargin = marginBottom.Value;
            if (marginLeft.HasValue)
                pageSetup.LeftMargin = marginLeft.Value;
            if (marginRight.HasValue)
                pageSetup.RightMargin = marginRight.Value;
            
            if (pageWidth.HasValue)
                pageSetup.PageWidth = pageWidth.Value;
            if (pageHeight.HasValue)
                pageSetup.PageHeight = pageHeight.Value;
            
            if (!string.IsNullOrEmpty(orientation))
            {
                pageSetup.Orientation = orientation.ToLower() switch
                {
                    "landscape" => Orientation.Landscape,
                    "portrait" => Orientation.Portrait,
                    _ => Orientation.Portrait
                };
            }
            
            if (headerDistance.HasValue)
                pageSetup.HeaderDistance = headerDistance.Value;
            if (footerDistance.HasValue)
                pageSetup.FooterDistance = footerDistance.Value;
        }

        doc.Save(outputPath);
        
        var result = $"Page setup updated successfully: {outputPath}\n";
        
        // Report what was changed
        if (marginTop.HasValue || marginBottom.HasValue || marginLeft.HasValue || marginRight.HasValue)
        {
            result += "Margins set to:\n";
            if (marginTop.HasValue) result += $"  Top: {marginTop.Value:F2} pt ({marginTop.Value / 28.35:F2} cm)\n";
            if (marginBottom.HasValue) result += $"  Bottom: {marginBottom.Value:F2} pt ({marginBottom.Value / 28.35:F2} cm)\n";
            if (marginLeft.HasValue) result += $"  Left: {marginLeft.Value:F2} pt ({marginLeft.Value / 28.35:F2} cm)\n";
            if (marginRight.HasValue) result += $"  Right: {marginRight.Value:F2} pt ({marginRight.Value / 28.35:F2} cm)\n";
        }
        
        if (pageWidth.HasValue || pageHeight.HasValue)
        {
            result += "Page size set to:\n";
            if (pageWidth.HasValue) result += $"  Width: {pageWidth.Value:F2} pt ({pageWidth.Value / 28.35:F2} cm)\n";
            if (pageHeight.HasValue) result += $"  Height: {pageHeight.Value:F2} pt ({pageHeight.Value / 28.35:F2} cm)\n";
        }
        
        if (!string.IsNullOrEmpty(orientation))
            result += $"Orientation: {orientation}\n";
        
        if (headerDistance.HasValue || footerDistance.HasValue)
        {
            result += "Header/Footer distance set to:\n";
            if (headerDistance.HasValue) result += $"  Header: {headerDistance.Value:F2} pt ({headerDistance.Value / 28.35:F2} cm)\n";
            if (footerDistance.HasValue) result += $"  Footer: {footerDistance.Value:F2} pt ({footerDistance.Value / 28.35:F2} cm)\n";
        }
            
        return await Task.FromResult(result);
    }
}

