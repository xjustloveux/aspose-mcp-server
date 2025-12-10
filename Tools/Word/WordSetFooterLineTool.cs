using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeMcpServer.Tools;

public class WordSetFooterLineTool : IAsposeTool
{
    public string Description => "Set line (border or shape) for footer in a Word document (fine-grained control)";

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
            lineStyle = new
            {
                type = "string",
                description = "Line style: border (paragraph border) or shape (graphic line, recommended)",
                @enum = new[] { "border", "shape" }
            },
            lineWidth = new
            {
                type = "number",
                description = "Line width in points (default: 1.0 for shape, 0.5 for border)"
            },
            lineColor = new
            {
                type = "string",
                description = "Line color (hex format, e.g., '000000' for black, default: black)"
            },
            position = new
            {
                type = "string",
                description = "Line position: above (above content) or below (below content, default)",
                @enum = new[] { "above", "below" }
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based). Default: 0. Use -1 to apply to all sections"
            },
            removeExisting = new
            {
                type = "boolean",
                description = "Remove existing line before adding new one (default: true)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var lineStyle = arguments?["lineStyle"]?.GetValue<string>() ?? "shape";
        var lineWidth = arguments?["lineWidth"]?.GetValue<double?>();
        var lineColor = arguments?["lineColor"]?.GetValue<string>() ?? "000000";
        var position = arguments?["position"]?.GetValue<string>() ?? "below";
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? 0;
        var removeExisting = arguments?["removeExisting"]?.GetValue<bool>() ?? true;

        var doc = new Document(path);
        
        var sections = sectionIndex == -1 ? doc.Sections.Cast<Section>() : new[] { doc.Sections[sectionIndex] };

        foreach (Section section in sections)
        {
            var footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
            if (footer == null)
            {
                footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
                section.HeadersFooters.Add(footer);
            }

            // Remove existing lines if requested
            if (removeExisting)
            {
                // Remove shape lines
                var shapes = footer.GetChildNodes(NodeType.Shape, true).Cast<Shape>()
                    .Where(s => s.ShapeType == ShapeType.Line).ToList();
                foreach (var shape in shapes)
                {
                    shape.Remove();
                }
                
                // Remove border from first paragraph
                var firstPara = footer.FirstParagraph;
                if (firstPara != null)
                {
                    firstPara.ParagraphFormat.Borders.Top.LineStyle = LineStyle.None;
                }
            }

            if (lineStyle == "shape")
            {
                // Add shape line
                var linePara = new Paragraph(doc);
                linePara.ParagraphFormat.SpaceBefore = 0;
                linePara.ParagraphFormat.SpaceAfter = 0;
                linePara.ParagraphFormat.LineSpacing = 1;
                linePara.ParagraphFormat.LineSpacingRule = LineSpacingRule.Exactly;
                
                var contentWidth = section.PageSetup.PageWidth - section.PageSetup.LeftMargin - section.PageSetup.RightMargin;
                var width = lineWidth ?? 1.0;
                
                var shape = new Shape(doc, ShapeType.Line);
                shape.Width = contentWidth;
                shape.Height = 0;
                shape.StrokeWeight = width;
                shape.StrokeColor = ParseColor(lineColor);
                shape.WrapType = WrapType.Inline;
                
                linePara.AppendChild(shape);
                
                // Insert based on position
                if (position == "above")
                {
                    footer.PrependChild(linePara);
                }
                else
                {
                    footer.AppendChild(linePara);
                }
            }
            else
            {
                // Add border line
                var firstPara = footer.FirstParagraph;
                if (firstPara == null)
                {
                    firstPara = new Paragraph(doc);
                    footer.AppendChild(firstPara);
                }
                
                var width = lineWidth ?? 0.5;
                firstPara.ParagraphFormat.Borders.Top.LineStyle = LineStyle.Single;
                firstPara.ParagraphFormat.Borders.Top.LineWidth = width;
                firstPara.ParagraphFormat.Borders.Top.Color = ParseColor(lineColor);
            }
        }

        doc.Save(outputPath);
        
        var sectionsDesc = sectionIndex == -1 ? "所有節" : $"第 {sectionIndex} 節";
        var styleDesc = lineStyle == "shape" ? "圖形線條" : "段落邊框";
        var positionDesc = position == "above" ? "上方" : "下方";
        
        return await Task.FromResult($"成功設定頁尾{styleDesc}（{positionDesc}）於 {sectionsDesc}");
    }
    
    private System.Drawing.Color ParseColor(string colorStr)
    {
        if (string.IsNullOrEmpty(colorStr))
            return System.Drawing.Color.Black;
        
        colorStr = colorStr.TrimStart('#');
        
        if (colorStr.Length == 6)
        {
            var r = Convert.ToInt32(colorStr.Substring(0, 2), 16);
            var g = Convert.ToInt32(colorStr.Substring(2, 2), 16);
            var b = Convert.ToInt32(colorStr.Substring(4, 2), 16);
            return System.Drawing.Color.FromArgb(r, g, b);
        }
        
        return System.Drawing.Color.Black;
    }
}

