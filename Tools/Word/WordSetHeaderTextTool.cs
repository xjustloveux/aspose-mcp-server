using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordSetHeaderTextTool : IAsposeTool
{
    public string Description => "Set header text content in a Word document (fine-grained control)";

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
            headerLeft = new
            {
                type = "string",
                description = "Header left section text (optional)"
            },
            headerCenter = new
            {
                type = "string",
                description = "Header center section text (optional)"
            },
            headerRight = new
            {
                type = "string",
                description = "Header right section text (optional)"
            },
            fontName = new
            {
                type = "string",
                description = "Font name (e.g., '標楷體', 'Arial'). If fontNameAscii and fontNameFarEast are provided, this will be used as fallback."
            },
            fontNameAscii = new
            {
                type = "string",
                description = "Font name for ASCII characters (English, e.g., 'Times New Roman')"
            },
            fontNameFarEast = new
            {
                type = "string",
                description = "Font name for Far East characters (Chinese/Japanese/Korean, e.g., '標楷體')"
            },
            fontSize = new
            {
                type = "number",
                description = "Font size in points (e.g., 10)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based). Default: 0. Use -1 to apply to all sections"
            },
            clearExisting = new
            {
                type = "boolean",
                description = "Clear existing header content before setting new content (default: true)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var headerLeft = arguments?["headerLeft"]?.GetValue<string>();
        var headerCenter = arguments?["headerCenter"]?.GetValue<string>();
        var headerRight = arguments?["headerRight"]?.GetValue<string>();
        var fontName = arguments?["fontName"]?.GetValue<string>();
        var fontNameAscii = arguments?["fontNameAscii"]?.GetValue<string>();
        var fontNameFarEast = arguments?["fontNameFarEast"]?.GetValue<string>();
        var fontSize = arguments?["fontSize"]?.GetValue<double?>();
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? 0;
        var clearExisting = arguments?["clearExisting"]?.GetValue<bool>() ?? true;

        var doc = new Document(path);
        
        bool hasContent = !string.IsNullOrEmpty(headerLeft) || !string.IsNullOrEmpty(headerCenter) || !string.IsNullOrEmpty(headerRight);
        if (!hasContent)
            return await Task.FromResult("警告：未提供任何頁首文字內容");

        var sections = sectionIndex == -1 ? doc.Sections.Cast<Section>() : new[] { doc.Sections[sectionIndex] };

        foreach (Section section in sections)
        {
            var header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
            if (header == null)
            {
                header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
                section.HeadersFooters.Add(header);
            }
            else if (clearExisting)
            {
                header.RemoveAllChildren();
            }

            bool useThreePart = !string.IsNullOrEmpty(headerLeft) || !string.IsNullOrEmpty(headerCenter) || !string.IsNullOrEmpty(headerRight);
            
            if (useThreePart)
            {
                var headerPara = new Paragraph(doc);
                
                // Add left text
                if (!string.IsNullOrEmpty(headerLeft))
                {
                    var leftRun = new Run(doc, headerLeft);
                    ApplyFontSettings(leftRun, fontName, fontNameAscii, fontNameFarEast, fontSize);
                    headerPara.AppendChild(leftRun);
                }
                
                // Add center text (with tab before it)
                if (!string.IsNullOrEmpty(headerCenter))
                {
                    headerPara.AppendChild(new Run(doc, "\t"));
                    var centerRun = new Run(doc, headerCenter);
                    ApplyFontSettings(centerRun, fontName, fontNameAscii, fontNameFarEast, fontSize);
                    headerPara.AppendChild(centerRun);
                }
                
                // Add right text (with tab before it)
                if (!string.IsNullOrEmpty(headerRight))
                {
                    headerPara.AppendChild(new Run(doc, "\t"));
                    var rightRun = new Run(doc, headerRight);
                    ApplyFontSettings(rightRun, fontName, fontNameAscii, fontNameFarEast, fontSize);
                    headerPara.AppendChild(rightRun);
                }
                
                header.AppendChild(headerPara);
            }
        }

        doc.Save(outputPath);
        
        var contentParts = new List<string>();
        if (!string.IsNullOrEmpty(headerLeft)) contentParts.Add("左");
        if (!string.IsNullOrEmpty(headerCenter)) contentParts.Add("中");
        if (!string.IsNullOrEmpty(headerRight)) contentParts.Add("右");
        
        var contentDesc = string.Join("、", contentParts);
        var sectionsDesc = sectionIndex == -1 ? "所有節" : $"第 {sectionIndex} 節";
        
        return await Task.FromResult($"成功設定頁首文字（{contentDesc}）於 {sectionsDesc}");
    }
    
    private void ApplyFontSettings(Run run, string? fontName, string? fontNameAscii, string? fontNameFarEast, double? fontSize)
    {
        // Set font names (priority: fontNameAscii/fontNameFarEast > fontName)
        if (!string.IsNullOrEmpty(fontNameAscii))
            run.Font.NameAscii = fontNameAscii;
        
        if (!string.IsNullOrEmpty(fontNameFarEast))
            run.Font.NameFarEast = fontNameFarEast;
        
        if (!string.IsNullOrEmpty(fontName))
        {
            // If fontNameAscii/FarEast are not set, use fontName for both
            if (string.IsNullOrEmpty(fontNameAscii) && string.IsNullOrEmpty(fontNameFarEast))
            {
                run.Font.Name = fontName;
            }
            else
            {
                // If only one is set, use fontName as fallback for the other
                if (string.IsNullOrEmpty(fontNameAscii))
                    run.Font.NameAscii = fontName;
                if (string.IsNullOrEmpty(fontNameFarEast))
                    run.Font.NameFarEast = fontName;
            }
        }
        
        if (fontSize.HasValue)
            run.Font.Size = fontSize.Value;
    }
}

