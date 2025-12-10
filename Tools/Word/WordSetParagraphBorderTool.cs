using System.Text.Json.Nodes;
using Aspose.Words;

namespace AsposeMcpServer.Tools;

public class WordSetParagraphBorderTool : IAsposeTool
{
    public string Description => "Set borders for a paragraph in a Word document";

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
            paragraphIndex = new
            {
                type = "number",
                description = "Paragraph index (0-based)"
            },
            sectionIndex = new
            {
                type = "number",
                description = "Section index (0-based). Default: 0"
            },
            // Border settings for each side
            borderTop = new
            {
                type = "boolean",
                description = "Show top border (default: false)"
            },
            borderBottom = new
            {
                type = "boolean",
                description = "Show bottom border (default: false)"
            },
            borderLeft = new
            {
                type = "boolean",
                description = "Show left border (default: false)"
            },
            borderRight = new
            {
                type = "boolean",
                description = "Show right border (default: false)"
            },
            // Border style (applies to all sides)
            lineStyle = new
            {
                type = "string",
                description = "Border line style: none, single, double, dotted, dashed, thick",
                @enum = new[] { "none", "single", "double", "dotted", "dashed", "thick" }
            },
            lineWidth = new
            {
                type = "number",
                description = "Border line width in points (default: 0.5)"
            },
            lineColor = new
            {
                type = "string",
                description = "Border line color (hex format, e.g., '000000' for black, default: black)"
            },
            // Individual side settings (optional, overrides common settings)
            topLineStyle = new
            {
                type = "string",
                description = "Top border line style (overrides lineStyle for top)"
            },
            topLineWidth = new
            {
                type = "number",
                description = "Top border line width in points"
            },
            topLineColor = new
            {
                type = "string",
                description = "Top border line color (hex format)"
            },
            bottomLineStyle = new
            {
                type = "string",
                description = "Bottom border line style (overrides lineStyle for bottom)"
            },
            bottomLineWidth = new
            {
                type = "number",
                description = "Bottom border line width in points"
            },
            bottomLineColor = new
            {
                type = "string",
                description = "Bottom border line color (hex format)"
            },
            leftLineStyle = new
            {
                type = "string",
                description = "Left border line style (overrides lineStyle for left)"
            },
            leftLineWidth = new
            {
                type = "number",
                description = "Left border line width in points"
            },
            leftLineColor = new
            {
                type = "string",
                description = "Left border line color (hex format)"
            },
            rightLineStyle = new
            {
                type = "string",
                description = "Right border line style (overrides lineStyle for right)"
            },
            rightLineWidth = new
            {
                type = "number",
                description = "Right border line width in points"
            },
            rightLineColor = new
            {
                type = "string",
                description = "Right border line color (hex format)"
            }
        },
        required = new[] { "path", "paragraphIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var paragraphIndex = arguments?["paragraphIndex"]?.GetValue<int>() ?? throw new ArgumentException("paragraphIndex is required");
        var sectionIndex = arguments?["sectionIndex"]?.GetValue<int>() ?? 0;

        var doc = new Document(path);
        
        if (sectionIndex >= doc.Sections.Count)
            throw new ArgumentException($"Section index {sectionIndex} out of range (total sections: {doc.Sections.Count})");
        
        var section = doc.Sections[sectionIndex];
        var paragraphs = section.Body.GetChildNodes(NodeType.Paragraph, true).Cast<Paragraph>().ToList();
        
        if (paragraphIndex >= paragraphs.Count)
            throw new ArgumentException($"Paragraph index {paragraphIndex} out of range (total paragraphs: {paragraphs.Count})");
        
        var para = paragraphs[paragraphIndex];
        var borders = para.ParagraphFormat.Borders;
        
        // Default values
        var defaultLineStyle = arguments?["lineStyle"]?.GetValue<string>() ?? "single";
        var defaultLineWidth = arguments?["lineWidth"]?.GetValue<double>() ?? 0.5;
        var defaultLineColor = arguments?["lineColor"]?.GetValue<string>() ?? "000000";
        
        // Top border
        if (arguments?["borderTop"]?.GetValue<bool>() == true)
        {
            var lineStyle = arguments?["topLineStyle"]?.GetValue<string>() ?? defaultLineStyle;
            var lineWidth = arguments?["topLineWidth"]?.GetValue<double>() ?? defaultLineWidth;
            var lineColor = arguments?["topLineColor"]?.GetValue<string>() ?? defaultLineColor;
            
            borders.Top.LineStyle = GetLineStyle(lineStyle);
            borders.Top.LineWidth = lineWidth;
            borders.Top.Color = ParseColor(lineColor);
        }
        else
        {
            borders.Top.LineStyle = LineStyle.None;
        }
        
        // Bottom border
        if (arguments?["borderBottom"]?.GetValue<bool>() == true)
        {
            var lineStyle = arguments?["bottomLineStyle"]?.GetValue<string>() ?? defaultLineStyle;
            var lineWidth = arguments?["bottomLineWidth"]?.GetValue<double>() ?? defaultLineWidth;
            var lineColor = arguments?["bottomLineColor"]?.GetValue<string>() ?? defaultLineColor;
            
            borders.Bottom.LineStyle = GetLineStyle(lineStyle);
            borders.Bottom.LineWidth = lineWidth;
            borders.Bottom.Color = ParseColor(lineColor);
        }
        else
        {
            borders.Bottom.LineStyle = LineStyle.None;
        }
        
        // Left border
        if (arguments?["borderLeft"]?.GetValue<bool>() == true)
        {
            var lineStyle = arguments?["leftLineStyle"]?.GetValue<string>() ?? defaultLineStyle;
            var lineWidth = arguments?["leftLineWidth"]?.GetValue<double>() ?? defaultLineWidth;
            var lineColor = arguments?["leftLineColor"]?.GetValue<string>() ?? defaultLineColor;
            
            borders.Left.LineStyle = GetLineStyle(lineStyle);
            borders.Left.LineWidth = lineWidth;
            borders.Left.Color = ParseColor(lineColor);
        }
        else
        {
            borders.Left.LineStyle = LineStyle.None;
        }
        
        // Right border
        if (arguments?["borderRight"]?.GetValue<bool>() == true)
        {
            var lineStyle = arguments?["rightLineStyle"]?.GetValue<string>() ?? defaultLineStyle;
            var lineWidth = arguments?["rightLineWidth"]?.GetValue<double>() ?? defaultLineWidth;
            var lineColor = arguments?["rightLineColor"]?.GetValue<string>() ?? defaultLineColor;
            
            borders.Right.LineStyle = GetLineStyle(lineStyle);
            borders.Right.LineWidth = lineWidth;
            borders.Right.Color = ParseColor(lineColor);
        }
        else
        {
            borders.Right.LineStyle = LineStyle.None;
        }
        
        doc.Save(outputPath);
        
        var enabledBorders = new List<string>();
        if (arguments?["borderTop"]?.GetValue<bool>() == true) enabledBorders.Add("上");
        if (arguments?["borderBottom"]?.GetValue<bool>() == true) enabledBorders.Add("下");
        if (arguments?["borderLeft"]?.GetValue<bool>() == true) enabledBorders.Add("左");
        if (arguments?["borderRight"]?.GetValue<bool>() == true) enabledBorders.Add("右");
        
        var bordersDesc = enabledBorders.Count > 0 ? string.Join("、", enabledBorders) : "無";
        
        return await Task.FromResult($"成功設定段落 {paragraphIndex} 的邊框：{bordersDesc}");
    }
    
    private LineStyle GetLineStyle(string style)
    {
        return style.ToLower() switch
        {
            "none" => LineStyle.None,
            "single" => LineStyle.Single,
            "double" => LineStyle.Double,
            "dotted" => LineStyle.Dot,
            "dashed" => LineStyle.Single, // Dash not available, use Single instead
            "thick" => LineStyle.Thick,
            _ => LineStyle.Single
        };
    }
    
    private System.Drawing.Color ParseColor(string colorStr)
    {
        if (string.IsNullOrEmpty(colorStr))
            return System.Drawing.Color.Black;
        
        // Remove # if present
        colorStr = colorStr.TrimStart('#');
        
        if (colorStr.Length == 6)
        {
            // RGB hex format
            var r = Convert.ToInt32(colorStr.Substring(0, 2), 16);
            var g = Convert.ToInt32(colorStr.Substring(2, 2), 16);
            var b = Convert.ToInt32(colorStr.Substring(4, 2), 16);
            return System.Drawing.Color.FromArgb(r, g, b);
        }
        
        return System.Drawing.Color.Black;
    }
}

