using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordSetHeaderTabStopsTool : IAsposeTool
{
    public string Description => "Set tab stops for header paragraphs in a Word document (fine-grained control)";

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
            tabStops = new
            {
                type = "array",
                description = "Tab stops array. Example: [{\"position\": 70.90, \"alignment\": \"Left\"}, {\"position\": 541.45, \"alignment\": \"Right\"}]. Set to empty array [] to remove all tab stops.",
                items = new
                {
                    type = "object",
                    properties = new
                    {
                        position = new { type = "number", description = "Tab stop position in points" },
                        alignment = new { type = "string", description = "Tab alignment: Left, Center, Right, Decimal, Bar", @enum = new[] { "Left", "Center", "Right", "Decimal", "Bar" } },
                        leader = new { type = "string", description = "Tab leader: None, Dots, Dashes, Line, Heavy, MiddleDot (default: None)", @enum = new[] { "None", "Dots", "Dashes", "Line", "Heavy", "MiddleDot" } }
                    }
                }
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based). Default: 0. Use -1 to apply to all sections"
            },
            paragraphIndex = new
            {
                type = "number",
                description = "Paragraph index in header (0-based). Default: 0 (first paragraph)"
            }
        },
        required = new[] { "path", "tabStops" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var tabStops = arguments?["tabStops"]?.AsArray() ?? throw new ArgumentException("tabStops is required");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? 0;
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int>() ?? 0;

        var doc = new Document(path);
        
        var sections = sectionIndex == -1 ? doc.Sections.Cast<Section>() : new[] { doc.Sections[sectionIndex] };

        foreach (Section section in sections)
        {
            var header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
            if (header == null)
            {
                // Create header if it doesn't exist
                header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
                section.HeadersFooters.Add(header);
            }

            var paragraphs = header.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
            
            // Ensure paragraph exists
            while (paragraphs.Count <= paragraphIndex)
            {
                var newPara = new Paragraph(doc);
                header.AppendChild(newPara);
                paragraphs.Add(newPara);
            }
            
            var para = paragraphs[paragraphIndex];
            var paraTabStops = para.ParagraphFormat.TabStops;
            
            // Clear existing tab stops
            paraTabStops.Clear();
            
            // Add new tab stops
            if (tabStops.Count > 0)
            {
                foreach (var tabStopJson in tabStops)
                {
                    var position = tabStopJson?["position"]?.GetValue<double>() ?? 0;
                    var alignmentStr = tabStopJson?["alignment"]?.GetValue<string>() ?? "Left";
                    var leaderStr = tabStopJson?["leader"]?.GetValue<string>() ?? "None";
                    
                    var alignment = alignmentStr switch
                    {
                        "Center" => TabAlignment.Center,
                        "Right" => TabAlignment.Right,
                        "Decimal" => TabAlignment.Decimal,
                        "Bar" => TabAlignment.Bar,
                        _ => TabAlignment.Left
                    };
                    
                    var leader = leaderStr switch
                    {
                        "Dots" => TabLeader.Dots,
                        "Dashes" => TabLeader.Dashes,
                        "Line" => TabLeader.Line,
                        "Heavy" => TabLeader.Heavy,
                        "MiddleDot" => TabLeader.MiddleDot,
                        _ => TabLeader.None
                    };
                    
                    paraTabStops.Add(new TabStop(position, alignment, leader));
                }
            }
        }

        doc.Save(outputPath);
        
        var tabCount = tabStops.Count;
        var sectionsDesc = sectionIndex == -1 ? "所有節" : $"第 {sectionIndex} 節";
        
        return await Task.FromResult($"成功設定頁首定位點（{tabCount} 個）於 {sectionsDesc}，段落 {paragraphIndex}");
    }
}

