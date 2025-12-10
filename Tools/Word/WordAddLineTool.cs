using System.Text.Json.Nodes;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace AsposeMcpServer.Tools;

public class WordAddLineTool : IAsposeTool
{
    public string Description => "Add a horizontal line to Word document (can be used in body, header, or footer)";

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
            location = new
            {
                type = "string",
                description = "Where to add the line: 'body' (current position in document body), 'header' (primary header), 'footer' (primary footer) (default: body)",
                @enum = new[] { "body", "header", "footer" }
            },
            position = new
            {
                type = "string",
                description = "Position in header/footer: 'start' (beginning), 'end' (end). For body, always inserts at current position. (default: end)",
                @enum = new[] { "start", "end" }
            },
            lineStyle = new
            {
                type = "string",
                description = "Line style: 'border' (paragraph border) or 'shape' (graphic line, recommended) (default: shape)",
                @enum = new[] { "border", "shape" }
            },
            lineWidth = new
            {
                type = "number",
                description = "Line width in points (default: 1.0)"
            },
            lineColor = new
            {
                type = "string",
                description = "Line color in hex (e.g., '000000' for black, default: '000000')"
            },
            width = new
            {
                type = "number",
                description = "Line length in points. If not specified, uses content area width (pageWidth - leftMargin - rightMargin)"
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var location = arguments?["location"]?.GetValue<string>() ?? "body";
        var position = arguments?["position"]?.GetValue<string>() ?? "end";
        var lineStyle = arguments?["lineStyle"]?.GetValue<string>() ?? "shape";
        var lineWidth = arguments?["lineWidth"]?.GetValue<double?>() ?? 1.0;
        var lineColor = arguments?["lineColor"]?.GetValue<string>() ?? "000000";
        var width = arguments?["width"]?.GetValue<double?>();

        var doc = new Document(path);
        var section = doc.FirstSection;

        // Calculate line width if not specified
        var calculatedWidth = width ?? (section.PageSetup.PageWidth - section.PageSetup.LeftMargin - section.PageSetup.RightMargin);

        Node? targetNode = null;
        string locationDesc = "";

        // Determine where to insert the line
        switch (location.ToLower())
        {
            case "header":
                var header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
                if (header == null)
                {
                    header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
                    section.HeadersFooters.Add(header);
                }
                targetNode = header;
                locationDesc = "header";
                break;

            case "footer":
                var footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
                if (footer == null)
                {
                    footer = new HeaderFooter(doc, HeaderFooterType.FooterPrimary);
                    section.HeadersFooters.Add(footer);
                }
                targetNode = footer;
                locationDesc = "footer";
                break;

            case "body":
            default:
                targetNode = section.Body;
                locationDesc = "document body";
                break;
        }

        if (targetNode == null)
            throw new InvalidOperationException($"Could not access {location}");

        // Create the line
        if (lineStyle == "shape")
        {
            // Shape line (recommended)
            var linePara = new Paragraph(doc);
            
            // CRITICAL: Set paragraph spacing to 0 to avoid blank lines
            linePara.ParagraphFormat.SpaceBefore = 0;
            linePara.ParagraphFormat.SpaceAfter = 0;
            linePara.ParagraphFormat.LineSpacing = 1; // Minimum line spacing
            linePara.ParagraphFormat.LineSpacingRule = LineSpacingRule.Exactly;
            
            var shape = new Shape(doc, ShapeType.Line);
            shape.Width = calculatedWidth;
            shape.Height = 0; // Horizontal line
            shape.StrokeWeight = lineWidth;
            shape.StrokeColor = ParseColor(lineColor);
            
            // Set shape to inline to avoid taking extra space
            shape.WrapType = WrapType.Inline;
            
            linePara.AppendChild(shape);

            // Insert based on position
            if (position == "start")
            {
                if (targetNode is HeaderFooter headerFooter)
                    headerFooter.PrependChild(linePara);
                else if (targetNode is Body body)
                    body.PrependChild(linePara);
            }
            else // end
            {
                if (targetNode is HeaderFooter headerFooter)
                    headerFooter.AppendChild(linePara);
                else if (targetNode is Body body)
                    body.AppendChild(linePara);
            }
        }
        else // border
        {
            // Paragraph border
            var linePara = new Paragraph(doc);
            linePara.ParagraphFormat.Borders.Bottom.LineStyle = LineStyle.Single;
            linePara.ParagraphFormat.Borders.Bottom.LineWidth = lineWidth;
            linePara.ParagraphFormat.Borders.Bottom.Color = ParseColor(lineColor);

            // Insert based on position
            if (position == "start")
            {
                if (targetNode is HeaderFooter headerFooter)
                    headerFooter.PrependChild(linePara);
                else if (targetNode is Body body)
                    body.PrependChild(linePara);
            }
            else // end
            {
                if (targetNode is HeaderFooter headerFooter)
                    headerFooter.AppendChild(linePara);
                else if (targetNode is Body body)
                    body.AppendChild(linePara);
            }
        }

        doc.Save(outputPath);

        return await Task.FromResult($"成功在 {locationDesc} 的 {position} 位置插入線條\n" +
                                      $"線條樣式: {lineStyle}\n" +
                                      $"線條寬度: {lineWidth} pt\n" +
                                      $"線條長度: {calculatedWidth:F2} pt ({calculatedWidth / 28.35:F2} cm)\n" +
                                      $"線條顏色: #{lineColor}");
    }

    private System.Drawing.Color ParseColor(string hexColor)
    {
        hexColor = hexColor.TrimStart('#');
        if (hexColor.Length == 6)
        {
            var r = Convert.ToByte(hexColor.Substring(0, 2), 16);
            var g = Convert.ToByte(hexColor.Substring(2, 2), 16);
            var b = Convert.ToByte(hexColor.Substring(4, 2), 16);
            return System.Drawing.Color.FromArgb(r, g, b);
        }
        return System.Drawing.Color.Black;
    }
}

